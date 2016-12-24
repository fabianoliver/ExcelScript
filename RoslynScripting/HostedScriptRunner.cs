using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Classification;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Emit;
using Microsoft.CodeAnalysis.Formatting;
using Microsoft.CodeAnalysis.Host.Mef;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Text;
using RoslynScripting.Internal;
using RoslynScripting.Internal.Marshalling;
using ScriptingAbstractions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RoslynScripting
{
    internal sealed class LastCompilationInfo
    {
        public int RecompilationTriggerHash { get; set; }
        public Assembly CompiledAssembly { get; set; }
        public MethodInfo EntryPoint { get; set; }
        public Document Document { get; set; }
    }

    /// <summary>
    /// Instances of this class will be created and executed in the correct AppDomain, i.e. the domain in which the script is intended to be run.
    /// </summary>
    internal sealed class HostedScriptRunner<TGlobals> : MarshalByRefObject, IScriptRunner
        where TGlobals : class, IScriptGlobals
    {
        internal Guid UniqueId { get; set; } = Guid.NewGuid();

        private LastCompilationInfo m_LastCompilation = null; // for caching results/compilations

        /// <summary>
        /// Runs the script asynchronously
        /// </summary>
        /// <param name="factory">Factory method to create an IScriptGlobals instance in the current (= the script's) AppDomain</param>
        /// <param name="parameters">Full list of IParameterValues to be passed to the script</param>
        /// <param name="scriptCode">Code of the script</param>
        /// <param name="Options">Options for script execution, compilation or invocation</param>
        /// <param name="submission">Returns an instance of the compiler-generated wrapped script class (e.g. an instance of Submission#0)</param>
        /// <returns>A RemoteTask, so the caller (which could reside in a different AppDomain) can unwrap the results asynchronously</returns>
        public RemoteTask<object> RunAsync(Func<AppDomain, IScriptGlobals> factory, IParameterValue[] parameters, string originalScriptCode, ScriptingOptions Options)
        {
            string debug_code_file_path = (Debugger.IsAttached) ? Path.Combine(Path.GetTempPath(), "ExcelScript.DebuggedCode.csx") : null;
            IScriptGlobals globalsInstance = null;

            try
            {
                var rewrittenScriptCode = RewriteCode(originalScriptCode, parameters.Select(x => x.Parameter), debug_code_file_path);

                if (debug_code_file_path != null)
                {
                    // todo: might cause problems in multi threaded scenarios / whenever we're running through the .RunAsync() function in parallel.
                    // possibly create varying or even fully random file paths - though though that might need some cleanup logic
                    string line_directive = (debug_code_file_path == null) ? String.Empty : "#line 1 \"" + debug_code_file_path + "\"" + Environment.NewLine;
                    File.WriteAllText(debug_code_file_path, rewrittenScriptCode);
                    rewrittenScriptCode = line_directive + rewrittenScriptCode;
                }

                MethodInfo entryMethod;
                var assembly = GetOrCreateScriptAssembly(parameters.Select(x => x.Parameter).ToArray(), originalScriptCode, rewrittenScriptCode, Options, debug_code_file_path, out entryMethod);
                globalsInstance = CreateGlobals(factory, parameters);
    
                var result = RemoteTask.ServerStart<object>(cts => InvokeFactoryAsync(entryMethod, globalsInstance));
                return result;
            }
            finally
            {
                if (!String.IsNullOrEmpty(debug_code_file_path) && File.Exists(debug_code_file_path))
                {
                    File.Delete(debug_code_file_path);
                }

                IDisposable globalsDisposable = globalsInstance as IDisposable;
                if(globalsDisposable != null)
                {
                    globalsDisposable.Dispose();
                }
            }
        }

        internal class Range
        {
            public ClassifiedSpan ClassifiedSpan { get; private set; }
            public string Text { get; private set; }

            public Range(string classification, Microsoft.CodeAnalysis.Text.TextSpan span, SourceText text) :
                this(classification, span, text.GetSubText(span).ToString())
            {
            }

            public Range(string classification, Microsoft.CodeAnalysis.Text.TextSpan span, string text) :
                this(new ClassifiedSpan(classification, span), text)
            {
            }

            public Range(ClassifiedSpan classifiedSpan, string text)
            {
                this.ClassifiedSpan = classifiedSpan;
                this.Text = text;
            }

            public string ClassificationType
            {
                get { return ClassifiedSpan.ClassificationType; }
            }

            public Microsoft.CodeAnalysis.Text.TextSpan TextSpan
            {
                get { return ClassifiedSpan.TextSpan; }
            }
        }

        public RemoteTask<string> GetRtfFormattedCodeAsync(string originalScriptCode, ScriptingOptions Options, FormatColorScheme ColorScheme)
        {
            var result = RemoteTask.ServerStart<string>(cts => InternalGetRtfFormattedCodeAsync(originalScriptCode, Options, ColorScheme));
            return result;
        }

        public RemoteTask<FormattedText> GetFormattedCodeAsync(string originalScriptCode, ScriptingOptions Options, FormatColorScheme ColorScheme)
        {
            var result = RemoteTask.ServerStart<FormattedText>(cts => InternalGetFormattedCodeAsync(originalScriptCode, Options, ColorScheme));
            return result;
        }

        internal async Task<FormattedText> InternalGetFormattedCodeAsync(string originalScriptCode, ScriptingOptions Options, FormatColorScheme ColorScheme)
        {
            var document = GetOrCreateDocument(Array.Empty<IParameter>(), originalScriptCode, Options);
            document = Formatter.FormatAsync(document).Result;
            SourceText documentText = await document.GetTextAsync();

            IEnumerable<ClassifiedSpan> classifiedSpans = await Classifier.GetClassifiedSpansAsync(document, Microsoft.CodeAnalysis.Text.TextSpan.FromBounds(0, documentText.Length));

            var ranges = classifiedSpans.Select(classifiedSpan =>
                new Range(classifiedSpan, documentText.GetSubText(classifiedSpan.TextSpan).ToString()));

            var LineSpans = documentText.Lines.Select(x => x.SpanIncludingLineBreak);
            var result = new FormattedText();
            
            var originalText = documentText.ToString();
            var codeBuilder = new StringBuilder();

            var segmentedSpans = DivideSpans(LineSpans, ranges.Select(x => x.TextSpan)).ToArray();

            FormattedTextLine line = result.AppendLine();

            foreach (var span in segmentedSpans)
            {
                var info = ranges.SingleOrDefault(x => x.TextSpan.CompareTo(span) == 0);

                TextFormat format;

                if (info != null)
                {
                    var color = ColorScheme.GetColorForKeyword(info.ClassificationType);
                    format = new TextFormat(color);
                } else
                {
                    var color = ColorScheme.Unknown;
                    format = new TextFormat(color);
                }

                var text = originalText.Substring(span.Start, span.Length);
            
                if (text.EndsWith(Environment.NewLine))
                {
                    text = text.Substring(0, text.Length - Environment.NewLine.Length);
                    line.AppendText(text, format);
                    line = result.AppendLine();
                } else
                {
                    line.AppendText(text, format);
                }
                    

                //{
                    //text = text.Substring(0, text.Length - Environment.NewLine.Length);
                    
                   // line = result.AppendLine();
              /*  } else
                {
                    line.AppendText(text, format);
                }*/
            }

            return result;
        }

        internal async Task<string> InternalGetRtfFormattedCodeAsync(string originalScriptCode, ScriptingOptions Options, FormatColorScheme ColorScheme)
        {

            var document = GetOrCreateDocument(Array.Empty<IParameter>(), originalScriptCode, Options);
            document = Formatter.FormatAsync(document).Result;
            SourceText documentText = await document.GetTextAsync();

            IEnumerable<ClassifiedSpan> classifiedSpans = await Classifier.GetClassifiedSpansAsync(document, Microsoft.CodeAnalysis.Text.TextSpan.FromBounds(0, documentText.Length));
   
            var ranges = classifiedSpans.Select(classifiedSpan =>
                new Range(classifiedSpan, documentText.GetSubText(classifiedSpan.TextSpan).ToString()));

            var LineSpans = documentText.Lines.Select(x => x.SpanIncludingLineBreak);


            var result = new StringBuilder();

            // RTF Header
            result.AppendLine(@"{\rtf1\ansi\deff0");

            // RTF Color table
            var colorDefinitions = ColorScheme.GetColorsInRtfOrder().Select(x => $"\\red{x.R}\\green{x.G}\\blue{x.B};");
            result.AppendLine(@"{\colortbl;" + String.Join(String.Empty, colorDefinitions) + "}");


            var originalText = documentText.ToString();
            var codeBuilder = new StringBuilder();

            var segmentedSpans = DivideSpans(LineSpans, ranges.Select(x => x.TextSpan)).ToArray();

            foreach(var span in segmentedSpans)
            {
                var info = ranges.SingleOrDefault(x => x.TextSpan.CompareTo(span) == 0);

                if(info != null)
                {
                    var colorIndex = ColorScheme.GetColorIndexForKeyword(info.ClassificationType);
                    codeBuilder.AppendLine(Environment.NewLine + @"\cf" + colorIndex);
                }

                var text = originalText.Substring(span.Start, span.Length);
                text = GetRtfUnicodeEscapedString(text);
                text = text.Replace(Environment.NewLine, @"\line" + Environment.NewLine);
                codeBuilder.Append(text);
            }

            result.Append(codeBuilder.ToString());
            result.AppendLine("}");

            return result.ToString();
        }

        /// <summary>
        /// Uses <paramref name="divisors"/> to segment <paramref name="source"/>
        /// E.g.
        /// source:  0 - 4 5  -  9
        /// divisor:  1-3      7-9
        /// result:  0-1 1-3 3-4 4-5 5-7 7-9
        /// </summary>
        /// <param name="source">A non-overlapping enumerable of TextSpans</param>
        /// <param name="divisors">A non-overlapping enumerable of TextSpans</param>
        /// <returns></returns>
        private IEnumerable<Microsoft.CodeAnalysis.Text.TextSpan> DivideSpans(IEnumerable<Microsoft.CodeAnalysis.Text.TextSpan> source, IEnumerable<Microsoft.CodeAnalysis.Text.TextSpan> divisors)
        {
            var orderedSource = source.OrderBy(x => x.Start);
            int previousEnd = 0;
            foreach (var sourceSpan in orderedSource)
            {
                var intersectingDivisors = divisors.Select(x => x.Intersection(sourceSpan)).Where(x => x != null).Select(x => (Microsoft.CodeAnalysis.Text.TextSpan)x).ToArray();
                
                while(previousEnd < sourceSpan.End)
                {
                    var nextDivisor = intersectingDivisors.Where(x => x.Start >= previousEnd).OrderBy(x => x.Start).FirstOrDefault();

                    if(nextDivisor == default(Microsoft.CodeAnalysis.Text.TextSpan))
                    {
                        // Return everything from i til end of sourceSpan
                        int start = previousEnd;
                        int end = sourceSpan.End;
                        previousEnd = end;
                        yield return Microsoft.CodeAnalysis.Text.TextSpan.FromBounds(start, end);
                    } else if(nextDivisor.Start > previousEnd)
                    {
                        // there is a gap between the current position and the next divisor
                        int start = previousEnd;
                        int end = nextDivisor.Start;
                        previousEnd = end;
                        yield return Microsoft.CodeAnalysis.Text.TextSpan.FromBounds(start, end);
                    } else if(nextDivisor.Start == previousEnd)
                    {
                        // there is a next divisor, and its position is NOW
                        var intersection = (Microsoft.CodeAnalysis.Text.TextSpan)nextDivisor.Intersection(sourceSpan);
                        previousEnd = intersection.End;
                        yield return intersection;
                    } else
                    {
                        throw new InvalidOperationException("Divisor out of bounds");
                    }

                }
            }
        }



        // From http://stackoverflow.com/questions/1368020/how-to-output-unicode-string-to-rtf-using-c
        private static string GetRtfUnicodeEscapedString(string s)
        {
            var sb = new StringBuilder();
            foreach (var c in s)
            {
                if (c <= 0x7f)
                    sb.Append(c);
                else
                    sb.Append("\\u" + Convert.ToUInt32(c) + "?");
            }
            return sb.ToString();
        }

        //[Serializable]
       // [System.Runtime.Serialization.KnownType(typeof(Parameter))]
        internal class ParseResult : MarshalByRefObject, IParseResult
        {
            public string RefactoredCode { get; private set; }
            public string EntryMethodName { get; private set; }

            private IParameter[] Parameters { get; set; }

            IEnumerable<IParameter> IParseResult.Parameters
            {
                get
                {
                    return Parameters;
                }
            }

            public ParseResult(string RefactoredCode, IParameter[] Parameters, string EntryMethodName)
            {
                this.RefactoredCode = RefactoredCode;
                this.Parameters = Parameters;
                this.EntryMethodName = EntryMethodName;
            }
        }

        public RemoteTask<IParseResult> ParseAsync(string scriptCode, ScriptingOptions Options, Func<MethodInfo[], MethodInfo> EntryMethodSelector, Func<MethodInfo, IParameter[]> EntryMethodParameterFactory)
        {
            MethodInfo entryMethod;
            Assembly assembly = GetOrCreateScriptAssembly(new IParameter[0], scriptCode, scriptCode, Options, null, out entryMethod);

            Type submissionType = entryMethod.DeclaringType;  // eg typeof(Submission#0)
            MethodInfo[] entryMethodCandidates = GetEntryMethodCandidatesFrom(submissionType);

            MethodInfo parsedEntryMethod = EntryMethodSelector(entryMethodCandidates);

            if (parsedEntryMethod == null)
                throw new InvalidOperationException("No entry method function when parsing script");

            IParameter[] parsedParameters = EntryMethodParameterFactory(parsedEntryMethod);

            string refactoredCode = RefactorCodeToCallEntryMethod(scriptCode, parsedEntryMethod, parsedParameters);

            var result = new ParseResult(refactoredCode, parsedParameters, parsedEntryMethod.Name);
            return RemoteTask.ServerStart<IParseResult>(cts =>Task.FromResult<IParseResult>(result));
        }

        private static MethodInfo[] GetEntryMethodCandidatesFrom(Type submissionType)
        {
            var result = submissionType
                .GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .Where(x => x.ReturnType != null)  // filter out void-methods
                .Where(x => x.DeclaringType == submissionType)  // filter out any function that wasn't declared on the type itself (e.g. inherited methods)
                .Where(x => !x.IsSpecialName)  // exclude special names such as compiler-generated functions (e.g. <Initialize>)
                .ToArray();
            return result;
        }

        private static string RefactorCodeToCallEntryMethod(string code, MethodInfo EntryMethod, IEnumerable<IParameter> Parameters)
        {
            SyntaxTree tree = CSharpSyntaxTree.ParseText(code, CSharpParseOptions.Default.WithKind(SourceCodeKind.Script));

            var strStatement = Environment.NewLine + "return " +EntryMethod.Name + "(" + String.Join(", ", Parameters.Select(x => x.Name)) + ");";
            var statement = SyntaxFactory.GlobalStatement(SyntaxFactory.ParseStatement(strStatement));

            var root = (CompilationUnitSyntax)tree.GetRoot();
            root = root.InsertNodesAfter(root.Members.Last(), new[] { statement });

            var refactoredCode = root.GetText().ToString();
            return refactoredCode;
        }


        /*

        /// <summary>
        /// Replaces calls to user-defined variables to a call to the IScriptGlobals-Parameters dictionary
        /// e.g.
        /// var something = myUserIntVar + 1234;
        /// ->
        /// var something = (int)Parameters["myUserIntVar"] + 1234;
        /// </summary>
        private class VariableReplacer : CSharpSyntaxRewriter
        {
            private readonly IEnumerable<IParameter> UserParameters;
            private readonly string[] ParamNames;

            public VariableReplacer(IEnumerable<IParameter> UserParameters)
            {
                this.UserParameters = UserParameters;
                this.ParamNames = UserParameters.Select(x => x.Name).ToArray();
            }

            public override SyntaxNode VisitIdentifierName(IdentifierNameSyntax node)
            {
                string name = node.Identifier.Text;

                if(ParamNames.Contains(name))
                {
                    var parameter = UserParameters.Single(x => x.Name == name);
                    var kind = node.Kind();

                    string realName = $"({parameter.Type.FullName})Parameters[\"{name}\"]";
                    var newNode = SyntaxFactory.ParseExpression(realName);
                    return newNode;
                } else
                {
                    return node;
                }
            }
        }

        /// <summary>
        /// Re-Writes the user code, e.g. by addin additional variable definitions (user-defined parameters)
        /// and adding a line compiler directive, if needed.
        /// </summary>
        private static string RewriteCode(string scriptCode, IEnumerable<IParameterValue> ParameterValues, string debug_code_file_path)
        {
            SyntaxTree tree = CSharpSyntaxTree.ParseText(scriptCode, CSharpParseOptions.Default.WithKind(SourceCodeKind.Script));
            var visitor = new VariableReplacer(ParameterValues.Select(x => x.Parameter));

            var rewrittenNode = visitor.Visit(tree.GetRoot());
        
            string rewrittenCode = rewrittenNode.ToString();
            return rewrittenCode;
        }

    */

        /// <summary>
        /// Re-Writes the user code, e.g. by adding additional variable definitions (user-defined parameters)
        /// and adding a line compiler directive, if needed.
        /// </summary>
        private static string RewriteCode(string scriptCode, IEnumerable<IParameter> ParameterValues, string debug_code_file_path)
        {
            SyntaxTree tree = CSharpSyntaxTree.ParseText(scriptCode, CSharpParseOptions.Default.WithKind(SourceCodeKind.Script));

            IEnumerable<MemberDeclarationSyntax> statements = ParameterValues
                .Select(x => $"{x.Type.FullName} {x.Name} = ({x.Type.FullName})Parameters[\"{x.Name}\"];{Environment.NewLine}")
                .Select(x => SyntaxFactory.GlobalStatement(SyntaxFactory.ParseStatement(x)));


            var root = (CompilationUnitSyntax)tree.GetRoot();

            // Write parameter initializations
            if (statements.Any())
            {

                /* #line hidden and #line default statements to exclude added lines from counting; seemed to bug out visual studio bugging oddly though
                var _lineHide = SyntaxFactory.ParseLeadingTrivia("#line hidden" + Environment.NewLine);
                var _lineDefault = SyntaxFactory.ParseLeadingTrivia(Environment.NewLine + "#line default" + Environment.NewLine);

                statements = statements
                    .Select((x, i) =>
                    {
                        if (i == 0)
                            x = x.WithLeadingTrivia(_lineHide);

                        if (i == statements.Count() - 1)
                            x = x.WithTrailingTrivia(_lineDefault);

                        return x;
                    }).ToArray();
                    */

                if (!root.Members.Any())
                    root = root.AddMembers(statements.ToArray());
                else
                {
                    root = root.InsertNodesBefore(root.Members.First(), statements);
                }
            }

            var code = root.GetText().ToString();
            return code;
        }


        /// <summary>
        /// Creates an instance of the script globals using the factory, and adds all $Parameters to
        /// the resulting IScriptGlobals.Parameters-dictionary
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown if Parameters is null</exception>
        private IScriptGlobals CreateGlobals(Func<AppDomain, IScriptGlobals> factory, IEnumerable<IParameterValue> Parameters)
        {
            if (Parameters == null)
                throw new ArgumentNullException("Parameters");

            if (factory == null)
                throw new NullReferenceException("HostedScriptRunner globals factory was null");

            var instance = factory(AppDomain.CurrentDomain);

            foreach (var parameter in Parameters)
            {
                var transferableValue = parameter.GetTransferableValue();
                var converter = parameter.GetTransferableValueToOriginalValueConverter();
                var value = converter(transferableValue);

                if(value != null && !parameter.Parameter.Type.IsAssignableFrom(value.GetType()))
                    throw new InvalidCastException($"Expected re-converter parameter value to be of type {parameter.Parameter.Type.Name}, but was {value.GetType().Name}");
   
                instance.Parameters.Add(parameter.Parameter.Name, value);
            }

            return instance;
        }

        /// <summary>
        /// Since we cannot emit/define the globals type dynamically, we need to use a trick to allow the user to directly use their specified parameters.
        /// IScriptGlobals contains a dictionary with key = name of a user-defined parameter and value = the runtime value of that parameter (parameter value with which the script is being called by the user).
        /// Say this dictionary contains only one item (user defined parameter), key = "testParameter1", value = (double)12345.
        /// then we'll add additional code at the beginning of the user script, going
        /// double testParameter1 = (double)Parameters["testParameter1"];
        /// The user can then use all his parameters by name directly in their script, without needing to resort toe the Parameters-dictionary.
        /// 
        /// This function creates the necessary variable definitions & assignments for this.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown if Parameters is null</exception>
        private static string CreateCodeDefinitions(IEnumerable<IParameterValue> Parameters)
        {
            if (Parameters == null)
                throw new ArgumentNullException("Parameters");

            //var definitions = Parameters.Select(x => $"{x.Parameter.Type.FullName} {x.Parameter.Name} = ({x.Parameter.Type.FullName})System.Convert.ChangeType(GlobalValues[\"{x.Parameter.Name}\"], typeof({x.Parameter.Type.FullName}));");
            var definitions = Parameters.Select(x => $"{x.Parameter.Type.FullName} {x.Parameter.Name} = ({x.Parameter.Type.FullName})Parameters[\"{x.Parameter.Name}\"];");
            var code = String.Join(Environment.NewLine, definitions);

            return code;
        }
        /*
        /// <summary>
        /// Converts ScriptingAbstractions.Scriptionoptions to Microsoft.CodeAnalysis.Scripting.ScriptOptions
        /// </summary>
        /// <param name="ScriptingOptions">null is explicitly allowed! ScriptOptions.Default will be used in this case.</param>
        internal static ScriptOptions ToScriptOptions(ScriptingOptions ScriptingOptions)
        {
            var options = ScriptOptions.Default;

            if (ScriptingOptions == null)
                return options;

            options = options.WithReferences(ScriptingOptions.References.ToArray());
            options = options.WithImports(ScriptingOptions.Imports.ToArray());

            return options;
        }*/
/*
        /// <summary>
        /// Crates the CompilationOptions with which the $compilation should be compiled
        /// </summary>
        private static CompilationOptions CreateOptionsFor(Compilation compilation)
        {
            var compilationOptions = compilation.Options
                .WithOutputKind(OutputKind.DynamicallyLinkedLibrary);

            if (Debugger.IsAttached)
            {
                compilationOptions = compilationOptions.WithOptimizationLevel(OptimizationLevel.Debug);
            }

            return compilationOptions;
        }
        */
        /// <summary>
        /// Compiles the given compilation to raw data
        /// </summary>
        /// <param name="compilation">Compilation for which to emit data</param>
        /// <param name="assemblyRawData">Raw bytes of the compiled assembly</param>
        /// <param name="pdbRawData">Raw bytes of the debug symbols. This is ONLY emitted if there is currently a debugger attached to the process, otherwisen null.</param>
        private void Compile(Compilation compilation, out byte[] assemblyRawData, out byte[] pdbRawData)
        {
            try
            {
                using (var assemblyStream = new MemoryStream())
                {
                    using (var symbolStream = (Debugger.IsAttached ? new MemoryStream() : null))
                    {
                        var emitOptions = new EmitOptions(false, DebugInformationFormat.PortablePdb);

                        var emitResult = compilation.Emit(assemblyStream, symbolStream, options: emitOptions);

                        if (!emitResult.Success)
                        {
                            var errors = string.Join(Environment.NewLine, emitResult.Diagnostics.Select(x => x));
                            throw new InvalidOperationException("Emit error: " + errors);
                        }

                        assemblyRawData = assemblyStream.ToArray();
                        pdbRawData = symbolStream?.ToArray();
                    }
                }
            } finally
            {
                GC.Collect();
            }
        }


        /// <summary>
        /// Gets an equivalent of Type.AssemblyQualifiedName for the given typeSymbol
        /// </summary>
        private static string ToFullyQualifiedTypeName(ITypeSymbol typeSymbol)
        {
            var assemblyName = typeSymbol.ContainingAssembly?.ToDisplayString();
            var format = new SymbolDisplayFormat(typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces);

            var result = typeSymbol.ToDisplayString(format) + (String.IsNullOrWhiteSpace(assemblyName) ? String.Empty : ", " + assemblyName);
            return result;
        }

        /// <summary>
        /// Gets an equivalent of Type.AssemblyQualifiedName for the given parameterSymbol
        /// </summary>
        private static string ToFullyQualifiedTypeName(IParameterSymbol parameterSymbol)
        {
            return ToFullyQualifiedTypeName(parameterSymbol.Type);
        }

        /// <summary>
        /// From all assemblies loaded in the currently active AppDomain, returns the one that has the same FullName as $name.FullName
        /// </summary>
        private static Assembly ResolveAssembly(AssemblyName name)
        {
            var result = AppDomain
                .CurrentDomain
                .GetAssemblies()
                .SingleOrDefault(assembly => assembly.GetName().FullName.Equals(name.FullName));

            return result;
        }

        private static Type ResolveType(Assembly assembly, string typeName, bool throwOnError)
        {
            if (assembly == null)
                return Type.GetType(typeName, throwOnError);
            else
                return assembly.GetType(typeName, throwOnError);
        }

        /// <summary>
        /// Gets the runtime type from the current AppDomain that corresponds to the namedTypeSymbol's defined type
        /// </summary>
        internal static Type ToType(INamedTypeSymbol namedTypeSymbol)
        {
            var typeName = ToFullyQualifiedTypeName(namedTypeSymbol);
            var result = Type.GetType(typeName, ResolveAssembly, ResolveType, true);  // namedTypeSymbol.CodeBase is null, so we will do the assembly (and type) resolving ourselves here
            return result;
        }

        /// <summary>
        /// Gets the runtime type from the current AppDomain that corresponds to the parameterSymbol's defined type
        /// </summary>
        internal static Type ToType(IParameterSymbol parameterSymbol)
        {
            var typeName = ToFullyQualifiedTypeName(parameterSymbol);
            Type result = Type.GetType(typeName, ResolveAssembly, ResolveType, true);
            return result;
        }

        /// <summary>
        /// Based on the given $entryMethodSymbol, returns the corresponding runtime MethodInfo from the given assembly
        /// </summary>
        private static MethodInfo GetEntryMethod(Assembly assembly, IMethodSymbol entryMethodSymbol)
        {
            if (assembly == null)
                throw new ArgumentNullException("assembly");

            if (entryMethodSymbol == null)
                throw new ArgumentNullException("entryMethodSymbol");

            if (!entryMethodSymbol.IsStatic)
                throw new ArgumentException($"Expected entryMethodSymbol to describe a static method; symbol was {entryMethodSymbol}", nameof(entryMethodSymbol));

            if (entryMethodSymbol.Parameters.Count() != 1)
                throw new ArgumentException($"Expected entryMethodSymbol to have 1 parameter, but had {entryMethodSymbol.Parameters.Count()}; symbol was {entryMethodSymbol}", nameof(entryMethodSymbol));

            Type constructorParameterType = ToType(entryMethodSymbol.Parameters.Single());
            if (constructorParameterType != typeof(object[]))
                throw new ArgumentException($"Expected entryMethodSymbol to have 1 parameter of type object[], but had 1 parameter of type {constructorParameterType.FullName}; symbol was {entryMethodSymbol}", nameof(entryMethodSymbol));

            var type = ToType(entryMethodSymbol.ContainingType);

            if (type == null)
                throw new InvalidOperationException($"Could not get type for {entryMethodSymbol.ContainingType}");

            string factoryName = entryMethodSymbol.Name; // <Factory>
            var method = type.GetMethod(factoryName, BindingFlags.Static | BindingFlags.Public);

            if (method == null)
                throw new InvalidOperationException($"Could not get MethodInfo for method {method} on type {type.FullName}");

            return method;
        }
        
        /// <summary>
        /// Invokes the factory method (as given by $factoryMethodInfo), and passes the given globalsInstance to this method.
        /// Returns a Task<object> representing the result of the factory method (= the result of the script), and passes the submission object
        /// (i.e. the instance of the underlying script class, eg Submission#0) to the out $submission parameter.
        /// </summary>
        private static async Task<object> InvokeFactoryAsync(MethodInfo factoryMethodInfo, IScriptGlobals globalsInstance)
        {
            // [in]  args[0] -> an instance of the globals variables
            // [out] args[1] -> the constructor of the compiler-generated class (Submission#0) will set this array index to a reference to an instance of the generated class (-> the instance of Submission#0)
            var args = new object[2] { globalsInstance, null }; // args[0]: globals, args[1]: will return submission object

            var task = (Task<object>)factoryMethodInfo.Invoke(null, new object[] { args });
            
            var taskResult = await task;
            var submission = args[1];

            var result = taskResult;
            return result;
        }

        private Assembly GetOrCreateScriptAssembly(IParameter[] parameters, string originalScriptCode, string rewrittenScriptCode, ScriptingOptions Options, string debug_code_file_path = null, bool reflectionOnlyLoad = false)
        {
            MethodInfo _unused;
            return GetOrCreateScriptAssembly(parameters, originalScriptCode, rewrittenScriptCode, Options, debug_code_file_path, out _unused);
        }

        /// <summary>
        /// Compiles an assembly of the script if needed, or returns a cached version if possible
        /// </summary>
        private Assembly GetOrCreateScriptAssembly(IParameter[] parameters, string originalScriptCode, string rewrittenScriptCode, ScriptingOptions Options, string debug_code_file_path, out MethodInfo entryMethod)
        {
            int new_hash;
            // todo: this always seems to suggest we need to recompile
            if (NeedsRecompilationFor(parameters, originalScriptCode, Options, out new_hash))
            {
                // TODO: Can I somehow do this by compiling a Document, i.e. without CSharpScript?
                // The entry methods will look a bit different, but the major issue is - how do I manage globals in a custom compiled assembly?
                var options = ToScriptOptions(Options);

                var script = CSharpScript.Create(rewrittenScriptCode, options: options, globalsType: typeof(TGlobals));

                var compilation = script.GetCompilation();
                var compilationOptions = CreateOptionsFor(compilation);
                compilation = compilation.WithOptions(compilationOptions);

                byte[] assemblyRawData;
                byte[] pdbRawData;

                /*
                var document = CreateDocument(parameters, rewrittenScriptCode, Options);
                var compilation = document.Project.GetCompilationAsync().Result;


                var script = CSharpScript.Create(rewrittenScriptCode, options: ToScriptOptions(Options), globalsType: typeof(TGlobals));
                var scriptCompilation = script.GetCompilation();

                var info = scriptCompilation.ScriptCompilationInfo;

      
                // compilation = compilation.WithScriptCompilationInfo(info);
                var ep = compilation.GetEntryPoint(CancellationToken.None);*/


                Compile(compilation, out assemblyRawData, out pdbRawData);

                Assembly assembly = Assembly.Load(assemblyRawData, pdbRawData);

                var entryPointSymbol = compilation.GetEntryPoint(CancellationToken.None);
                entryMethod = GetEntryMethod(assembly, entryPointSymbol);
                this.m_LastCompilation = new LastCompilationInfo { CompiledAssembly = assembly, RecompilationTriggerHash = new_hash, EntryPoint = entryMethod };  // update cache info

                return assembly;
            }
            else
            {
                var compilation = m_LastCompilation.CompiledAssembly;
                entryMethod = m_LastCompilation.EntryPoint;
                return compilation;
            }
        }

        /// <summary>
        /// Crates the CompilationOptions with which the $compilation should be compiled
        /// </summary>
        private static CompilationOptions CreateOptionsFor(Compilation compilation)
        {
            var compilationOptions = compilation.Options
                .WithOutputKind(OutputKind.DynamicallyLinkedLibrary);

            if (Debugger.IsAttached)
            {
                compilationOptions = compilationOptions.WithOptimizationLevel(OptimizationLevel.Debug);
            }

            return compilationOptions;
        }

     

        internal static ScriptOptions ToScriptOptions(ScriptingOptions ScriptingOptions)
        {
            var options = ScriptOptions.Default;

            if (ScriptingOptions == null)
                return options;

            options = options.WithReferences(ScriptingOptions.References.ToArray());
            options = options.WithImports(ScriptingOptions.Imports.ToArray());

            return options;
        }


        private Document GetOrCreateDocument(IParameter[] parameters, string originalScriptCode, ScriptingOptions Options)
        {
            var rewrittenScriptCode = RewriteCode(originalScriptCode, parameters, null);

            int new_hash;
            // todo...

            if (NeedsRecompilationFor(parameters, originalScriptCode, Options, out new_hash))
            {
                return CreateDocument(parameters, rewrittenScriptCode, Options);
            } else
            {
                return m_LastCompilation.Document;
            }
        }

        private Document CreateDocument(IParameter[] parameters, string rewrittenScriptCode, ScriptingOptions Options)
        {
            var compositionContext = new System.Composition.Hosting.ContainerConfiguration()
                  .WithAssemblies(MefHostServices.DefaultAssemblies.Concat(new[] {
                                Assembly.Load("Microsoft.CodeAnalysis.Features"),
                                Assembly.Load("Microsoft.CodeAnalysis.CSharp.Features") }))
                  .CreateContainer();

            var workspace = new AdhocWorkspace(MefHostServices.Create(compositionContext));

            var metadataReferences = Options
                .References
                .Select(x => x.Location)
                .Select(x => MetadataReference.CreateFromFile(x))
                .ToArray();

            var compilationOptions = new CSharpCompilationOptions(
                OutputKind.DynamicallyLinkedLibrary,
                usings: Options.Imports,   
                optimizationLevel: (Debugger.IsAttached) ? OptimizationLevel.Debug : OptimizationLevel.Release);


            Solution solution = workspace.CurrentSolution;
            var project = solution.AddProject("SCRIPT-PROJECT", "SCRIPT-ASSEMBLY", LanguageNames.CSharp)
                .WithParseOptions(new CSharpParseOptions(kind: Microsoft.CodeAnalysis.SourceCodeKind.Script))
                .WithMetadataReferences(metadataReferences)
                .WithCompilationOptions(compilationOptions);

            Document document = project.AddDocument("SCRIPT-TEMP-DOCUMENT.cs", rewrittenScriptCode);
            return document;
        }

        /// <summary>
        /// Checks if we need to recompile the Assembly based on any changes to the script (i.e. its code, its parameters or its options)
        /// </summary>
        /// <param name="hash">RecompilationTriggerHash for the other input parameters ($parameters, $scriptCode and $Options)</param>
        /// <returns></returns>
        private bool NeedsRecompilationFor(IParameter[] parameters, string unmodifiedScriptCode, ScriptingOptions Options, out int hash)
        {
            hash = GetRecompilationTriggerHash(parameters, unmodifiedScriptCode, Options);

            if (m_LastCompilation == null)  // nothing has been compiled at all yet, so definitely need to compile now
                return true;
            else // check if last compilations hash was the same or not
                return hash != m_LastCompilation.RecompilationTriggerHash;
        }

        public bool NeedsRecompilationFor(IParameter[] parameters, string unmodifiedScriptCode, ScriptingOptions Options)
        {
            int _unused;
            return NeedsRecompilationFor(parameters, unmodifiedScriptCode, Options, out _unused);
        }


        private int GetRecompilationTriggerHash(IParameter[] parameters, string unmodifiedScriptCode, ScriptingOptions Options)
        {
            unchecked
            {
                int hash = (int)2166136261;

                hash = (hash * 16777619) ^ GetRecompilationTriggerHash(parameters);
                hash = (hash * 16777619) ^ ((unmodifiedScriptCode == null) ? 1 : unmodifiedScriptCode.GetHashCode());
                hash = (hash * 16777619) ^ ((Options == null) ? 1 : Options.GetHashCode());

                return hash;
            }
        }

        private static int GetRecompilationTriggerHash(IParameter[] parameters)
        {
            unchecked
            {
                int hash = (int)2166136261;

                foreach (var item in parameters)
                    hash = (hash * 16777619) ^ item.GetHashCode();

                return hash;
            }
        }
    }
}
