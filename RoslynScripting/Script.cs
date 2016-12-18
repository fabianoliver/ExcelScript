using ExcelScript.CommonUtilities;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Formatting;
using Microsoft.CodeAnalysis.Host.Mef;
using RoslynScripting.Internal;
using RoslynScripting.Internal.Marshalling;
using ScriptingAbstractions;
using ScriptingAbstractions.Factory;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Lifetime;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace RoslynScripting
{
    internal class HostingHelper
    {
        public static readonly ClientSponsor ClientSponsor = new ClientSponsor(TimeSpan.FromMinutes(5));
        private static int i = 0;
        private int ThisId = ++i;
        private static AppDomain _SharedScriptDomain = AppDomain.CreateDomain("SharedSandbox", AppDomain.CurrentDomain.Evidence, AppDomain.CurrentDomain.SetupInformation, AppDomain.CurrentDomain.PermissionSet);

        private Lazy<AppDomain> _IndividualDomain;

        public readonly AppDomain GlobalDomain;
        public readonly AppDomain SharedScriptDomain;
        public AppDomain IndividualScriptDomain
        {
            get
            {
                return _IndividualDomain.Value;
            }
        }

        public HostingHelper()
        {
            this.GlobalDomain = AppDomain.CurrentDomain;
            this.SharedScriptDomain = _SharedScriptDomain;
            this._IndividualDomain = new Lazy<AppDomain>(CreateIndividualDomain);
        }

        private AppDomain CreateIndividualDomain()
        {
            var domain = AppDomain.CreateDomain("IndividualSandbox" + GetUniqueDomainId(), AppDomain.CurrentDomain.Evidence, AppDomain.CurrentDomain.SetupInformation, AppDomain.CurrentDomain.PermissionSet);
            return domain;
        }

        public int GetUniqueDomainId()
        {
            return ThisId;
        }
    }

    internal class ScriptRunnerCacheInfo<TGlobals>
        where TGlobals : class, IScriptGlobals
    {
        public HostedScriptRunner<TGlobals> ScriptRunner { get; set; }
        public AppDomain HostingDomain { get; set; }
    }

    public class Script<TGlobals> : IScript<TGlobals>, IXmlSerializable, IDisposable
        where TGlobals : class, IScriptGlobals
    {
        /// <summary>
        /// Null in case of void
        /// </summary>
        [HashValue]
        public Type ReturnType { get; set; } = null;

        [HashEnumerable]
        public IList<IParameter> Parameters { get; protected set; } = new List<IParameter>();

        [HashValue]
        public string Code { get; set; } = String.Empty;

        [HashValue]
        public string Description { get; set; } = String.Empty;

        public readonly Guid UniqueId = Guid.NewGuid();  // just to make debugging a little easier

        [HashValue]
        private ScriptingOptions m_ScriptingOptions { get; set; }

        private readonly HostingHelper HostingHelper;
        private readonly ClientSponsor clientSponsor = new ClientSponsor();
        private ScriptRunnerCacheInfo<TGlobals> m_ScriptRunnerCacheInfo;
        private IDisposable m_appDomainDisposable;
        
        internal Func<AppDomain, IScriptGlobals> GlobalsFactory { private get; set; } // a factory that creates IScriptGlobals in the correct AppDomain

        /// <summary>
        /// Needed for XML Serialization
        /// </summary>
        private Script()
        {
            this.HostingHelper = new HostingHelper();
        }

        public Script(Func<AppDomain, IScriptGlobals> GlobalsFactory, ScriptingOptions ScriptingOptions)
            : this()
        {
            this.GlobalsFactory = GlobalsFactory;
            this.m_ScriptingOptions = ScriptingOptions;
        }

        public void Dispose()
        {
            TryUnregisterLease();

            if(m_appDomainDisposable != null)
            {
                m_appDomainDisposable.Dispose();
                m_appDomainDisposable = null;
            }
        }

        public static Task<string> GetRtfFormattedCodeAsync(string code, ScriptingOptions ScriptingOptions)
        {
            return GetRtfFormattedCodeAsync(code, ScriptingOptions, FormatColorScheme.LightTheme);
        }

        public static async Task<string> GetRtfFormattedCodeAsync(string code, ScriptingOptions ScriptingOptions, FormatColorScheme ColorScheme)
        {
            var script = new Script<TGlobals>(x => null, ScriptingOptions);
            script.Code = code;
            script.m_ScriptingOptions = ScriptingOptions;

            var hostedScriptRunner = script.GetOrCreateScriptRunner(Array.Empty<IParameter>(), script.Code, script.m_ScriptingOptions);
            var task = RemoteTask.ClientComplete<string>(hostedScriptRunner.GetRtfFormattedCodeAsync(code, ScriptingOptions, ColorScheme), CancellationToken.None);
            var result = await task;

            return result;
        }

        public static Task<FormattedText> GetFormattedCodeAsync(string code, ScriptingOptions ScriptingOptions)
        {
            return GetFormattedCodeAsync(code, ScriptingOptions, FormatColorScheme.LightTheme);
        }

        public static async Task<FormattedText> GetFormattedCodeAsync(string code, ScriptingOptions ScriptingOptions, FormatColorScheme ColorScheme)
        {
            var script = new Script<TGlobals>(x => null, ScriptingOptions);
            script.Code = code;
            script.m_ScriptingOptions = ScriptingOptions;

            var hostedScriptRunner = script.GetOrCreateScriptRunner(Array.Empty<IParameter>(), script.Code, script.m_ScriptingOptions);
            var task = RemoteTask.ClientComplete<FormattedText>(hostedScriptRunner.GetFormattedCodeAsync(code, ScriptingOptions, ColorScheme), CancellationToken.None);
            var result = await task;

            return result;
        }

        internal static async Task<IParseResult<TGlobals>> ParseFromAsync(Func<AppDomain, IScriptGlobals> GlobalsFactory, ScriptingOptions ScriptingOptions, string code, Func<MethodInfo[], MethodInfo> EntryMethodSelector, Func<MethodInfo, IParameter[]> EntryMethodParameterFactory)
        {
            var script = new Script<TGlobals>(GlobalsFactory, ScriptingOptions);
            script.Code = code;
            script.m_ScriptingOptions = ScriptingOptions;

            var hostedScriptRunner = script.GetOrCreateScriptRunner(new IParameter[0], script.Code, script.m_ScriptingOptions);
            var task = RemoteTask.ClientComplete<RoslynScripting.Internal.IParseResult>(hostedScriptRunner.ParseAsync(script.Code, script.m_ScriptingOptions, EntryMethodSelector, EntryMethodParameterFactory), CancellationToken.None);
            var parseResult = await task;

            script.Code = parseResult.RefactoredCode;
            script.Parameters = parseResult.Parameters.ToList();

            var result = new ParseResult<TGlobals>(script, parseResult.EntryMethodName);
            return result;
        }
        
        public IScriptRunResult Run(IEnumerable<IParameterValue> Parameters)
        {
            return RunAsync(Parameters).Result;
        }

        /// <exception cref="ArgumentNullException">Thrown if Parameters is null</exception>
        public async Task<IScriptRunResult> RunAsync(IEnumerable<IParameterValue> Parameters)
        {
            if (Parameters == null)
                throw new ArgumentNullException("Parameters");
            try
            {
                Parameters = CompleteParameterValues(Parameters);
            }
            catch (AggregateException ex)
            {
                return ScriptRunResult.Failure(ex);
            }
            
            var result = await InternalRunAsync(Parameters);
         
            Type resultType = result?.GetType();

            if (ReturnType == null)
                return ScriptRunResult.Success();

            if (!AreCompatible(ReturnType, resultType))
                return ScriptRunResult.Failure(new InvalidCastException($"Expected return type of script to be {ReturnType.Name}, but was {resultType.Name}"));

            return ScriptRunResult.Success(result);
        }

        /// <exception cref="ArgumentNullException">Thrown if Parameters is null</exception>
        private async Task<object> InternalRunAsync(IEnumerable<IParameterValue> Parameters)
        {
            if (Parameters == null)
                throw new ArgumentNullException("Parameters");

            var hostedScriptRunner = GetOrCreateScriptRunner(Parameters.Select(x => x.Parameter).ToArray(), this.Code, this.m_ScriptingOptions);
            var task = RemoteTask.ClientComplete<object>(hostedScriptRunner.RunAsync(this.GlobalsFactory, Parameters.ToArray(), this.Code, this.m_ScriptingOptions), CancellationToken.None);
            var result = await task;

            return result;
        }

        private IScriptRunner GetOrCreateScriptRunner(IParameter[] parameters, string scriptCode, ScriptingOptions Options)
        {
            var hostingType = (Options == null) ? HostingType.SharedSandboxAppDomain : Options.HostingType;
            AppDomain sandbox = GetOrCreateHostingDomain(hostingType);

            // We can return a cached script runner if 
            // A) there is one (duh...), AND
            // B) It does not need to recompile for the given input arguments, AND
            // C) the AppDomain in which the script runner is hosted is the same that we would like to use now
            if (m_ScriptRunnerCacheInfo != null && !m_ScriptRunnerCacheInfo.ScriptRunner.NeedsRecompilationFor(parameters, scriptCode, Options) && m_ScriptRunnerCacheInfo.HostingDomain == sandbox)
            {
                return m_ScriptRunnerCacheInfo.ScriptRunner;
            }
            else
            {
                TryUnregisterLease();

                var scriptRunner = (HostedScriptRunner<TGlobals>)Activator.CreateInstance(sandbox, typeof(HostedScriptRunner<TGlobals>).Assembly.FullName, typeof(HostedScriptRunner<TGlobals>).FullName).Unwrap();
                HostingHelper.ClientSponsor.Register(scriptRunner);
                
                this.m_ScriptRunnerCacheInfo = new ScriptRunnerCacheInfo<TGlobals> { HostingDomain = sandbox, ScriptRunner = scriptRunner };
                return scriptRunner;
            }
        }

        /// <summary>
        /// Tries to unregister the lease of the cached script runner
        /// </summary>
        private void TryUnregisterLease()
        {
            if (m_ScriptRunnerCacheInfo != null && m_ScriptRunnerCacheInfo.ScriptRunner != null)
            {
                HostingHelper.ClientSponsor.Unregister(m_ScriptRunnerCacheInfo.ScriptRunner);
                m_ScriptRunnerCacheInfo = null;
            }
        }

        /// <summary>
        /// Returns the AppDomain that shall be used to host script execution.
        /// </summary>
        /// <returns></returns>
        private AppDomain GetOrCreateHostingDomain(HostingType hostingType) {
            AppDomain domain;

            switch(hostingType)
            {
                case HostingType.GlobalAppDomain:
                    Trace.WriteLine($"Script {UniqueId} - executing in global domain");
                    domain = this.HostingHelper.GlobalDomain;
                    this.m_appDomainDisposable = null;
                    return domain;
                case HostingType.SharedSandboxAppDomain:
                    Trace.WriteLine($"Script {UniqueId} - executing in shared domain");
                    this.m_appDomainDisposable = null;
                    domain = this.HostingHelper.SharedScriptDomain;
                    return domain;
                case HostingType.IndividualScriptAppDomain:
                    Trace.WriteLine($"Script {UniqueId} - executing in individual domain");
                    domain = this.HostingHelper.IndividualScriptDomain;
                    this.m_appDomainDisposable = new DelegateDisposable(() => DisposeDomain(domain));
                    return domain;
                default:
                    throw new InvalidOperationException($"Unknown HostingType {hostingType}");
            }
        }
        
        private void DisposeDomain(AppDomain domain)
        {
            Trace.WriteLine($"Script {UniqueId} - Unloading AppDomain {domain}");
            AppDomain.Unload(domain);
        }
       
        /// <exception cref="ArgumentNullException">Thrown if expectedType or givenType is null</exception>
        private static bool AreCompatible(Type expectedType, Type givenType)
        {
            if (expectedType == null)
                throw new ArgumentNullException("expectedType");

            if (!expectedType.IsValueType && givenType == null)
                // expected type is a reference type and giventype = null, i.e. given type could be anything, including expectedtype - return true!
                return true;

            if (givenType == null)
                throw new ArgumentNullException("givenType");

            bool success = expectedType.IsAssignableFrom(givenType);

            bool assemblyEqual = expectedType.Assembly.Equals(givenType.Assembly);
            bool assemblyLocationEqual = expectedType.Assembly.Location.Equals(givenType.Assembly.Location);

            if (expectedType.AssemblyQualifiedName == givenType.AssemblyQualifiedName && (!assemblyEqual || !assemblyLocationEqual))
                throw new InvalidProgramException($"Types {expectedType.AssemblyQualifiedName} vs {givenType.AssemblyQualifiedName} appear to be loaded in different AppDomains");

            return success;
        }

        /// <exception cref="AggregateException">Thrown if not all non-optional parameters were given, or GivenParameters contains any parameters that were not part of the script</exception>
        /// <exception cref="ArgumentNullException">Thrown if GivenParameters is null</exception>
        private IEnumerable<IParameterValue> CompleteParameterValues(IEnumerable<IParameterValue> GivenParameters)
        {
            if (GivenParameters == null)
                throw new ArgumentNullException("GivenParameters");

            var givenParametersDefinitions = GivenParameters.Select(x => x.Parameter);

            var invalidParameters = givenParametersDefinitions.Except(Parameters);

            if(invalidParameters.Any())
            {
                var exceptions = invalidParameters.Select(x => new ArgumentException(String.Format("Parameter '{0}' is not a parameter of the script", x.Name), x.Name));
                throw new AggregateException("The given parameters contain parameters that are not aprt of the script definition", exceptions);
            }

            var mandatoryParameters = Parameters.Where(x => !x.IsOptional);
            var missingMandatoryParameters = mandatoryParameters.Except(givenParametersDefinitions);

            if(missingMandatoryParameters.Any())
            {
                var exceptions = missingMandatoryParameters.Select(x => new ArgumentException(String.Format("Parameter '{0}' was not given", x.Name), x.Name));
                throw new AggregateException("Not all required parameters were given", exceptions);
            }

            var optionalParameters = Parameters.Where(x => x.IsOptional);
            var missingOptionalParameters = optionalParameters.Except(givenParametersDefinitions);
            var missingOptionalParameterValues = missingMandatoryParameters.Select(x => (IParameterValue)new ParameterValue(x, x.DefaultValue));

            var allParameterValues = GivenParameters.Concat(missingOptionalParameterValues);
            return allParameterValues;
        }


        public override int GetHashCode()
        {
            int hash = HashHelper.HashOfAnnotated<Script<TGlobals>>(this);
            int domainId;
            
            // when e.g. two scripts are hosted with the individual domain setting, and everything else is equal (code, parameters etc)
            // we still don't want to return the same hash code, because they should be run in different domains.
            // hence, we acount for that case here:
            switch(this.m_ScriptingOptions.HostingType)
            {
                case HostingType.GlobalAppDomain:
                    domainId = -2;
                    break;
                case HostingType.SharedSandboxAppDomain:
                    domainId = -3;
                    break;
                case HostingType.IndividualScriptAppDomain:
                    domainId = this.HostingHelper.GetUniqueDomainId();
                    break;
                default:
                    throw new InvalidOperationException($"Unknown hosting type {m_ScriptingOptions.HostingType}");
            }

            unchecked
            {
                hash = (hash * 7) ^ domainId.GetHashCode();
            }

            return hash;
        }


        #region IXmlSerializable

        public XmlSchema GetSchema()
        {
            return null;
        }

        public void WriteXml(XmlWriter writer)
        {
            writer.Write(nameof(ReturnType), ReturnType);
            writer.WriteEnumerable(nameof(Parameters), Parameters);
            writer.Write(nameof(Code), Code);
            writer.Write(nameof(Description), Description);
            writer.Write(nameof(m_ScriptingOptions), m_ScriptingOptions);
        }

        public void ReadXml(XmlReader reader)
        {
            reader.MoveToCustomStart();

            Type _ReturnType;
            IEnumerable<IParameter> _Parameters;
            string _Code;
            string _Description;
            object _ScriptingOptions;

            reader.Read(nameof(ReturnType), out _ReturnType);
            _Parameters = reader.ReadEnumerable(nameof(Parameters)).Cast<IParameter>();
            reader.Read(nameof(Code), out _Code);
            reader.Read(nameof(Description), out _Description);
            reader.Read(nameof(m_ScriptingOptions), out _ScriptingOptions);

            this.ReturnType = _ReturnType;
            this.Parameters = _Parameters.ToList();
            this.Code = _Code;
            this.Description = _Description;
            this.m_ScriptingOptions = (ScriptingOptions)_ScriptingOptions;

            reader.ReadEndElement();
        }

        #endregion


    }
}
