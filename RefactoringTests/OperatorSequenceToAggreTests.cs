using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelAddIn3.Refactorings;
using Infotron.Parsing;

namespace RefactoringTests
{
    [TestClass]
    public class OperatorSequenceToAggregateTests
    {
        private readonly INodeRefactoring testee = new OperatorSequenceToAggregate();

        [TestMethod]
        public void TestPlus()
        {
            test("1+1+1", "SUM(1,1,1)");
        }

        private void cannotRefactor(string f)
        {
            Assert.IsFalse(testee.CanRefactor(f.Parse()), "Cannot refactor '{0}' but CanRefactor returned true", f);
        }

        private void test(string or, string target)
        {
            var orParsed = or.Parse();
            var targetParsed = target.Parse();

            Assert.IsTrue(testee.CanRefactor(orParsed), "Should be able to refactor '{0}'", or);

            var refactored = testee.Refactor(orParsed);

            Assert.AreEqual(Context.Empty.ProvideContext(targetParsed), Context.Empty.ProvideContext(refactored));
        }
    }
}
