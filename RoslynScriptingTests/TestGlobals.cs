using ScriptingAbstractions;
using System;
using System.Collections.Generic;

namespace RoslynScriptingTests
{
    public class TestGlobals : IScriptGlobals
    {
        public IDictionary<string, object> Parameters { get; private set; }
    }

    public class TestGlobalsFactory : IGlobalsFactory<TestGlobals>
    {
        object IGlobalsFactory.Create(AppDomain ExecutingDomain)
        {
            return Create(ExecutingDomain);
        }

        public TestGlobals Create(AppDomain ExecutingDomain)
        {
            return new TestGlobals();
        }
    }
}


