using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using BumbleBee.Refactorings;
using BumbleBee.Refactorings.Util;
using Infotron.Parsing;

namespace RefactoringTests
{
    [TestClass]
    public class OpToAggregateTests
    {
        private readonly IFormulaRefactoring testee = new OpToAggregate();

        [TestMethod]
        public void TestPlus01()
        {
            test("1+2", "SUM(1,2)");
        }

        [TestMethod]
        public void TestPlus02()
        {
            test("1+2+3", "SUM(1,2,3)");
        }

        [TestMethod]
        public void TestPlus03(){
            test("1+2*3+4*5", "SUM(1,2*3,4*5)");
        }

        [TestMethod]
        public void TestPlus04()
        {
            test("A1+A2", "SUM(A1,A2)");
        }

        [TestMethod]
        public void TestPlusCannot01()
        {
            cannotRefactor("+1");
        }

        [TestMethod]
        public void TestPlusCannot02()
        {
            cannotRefactor("1-1");
        }

        [TestMethod]
        public void TestMult()
        {
            test("1*2*3*4", "PRODUCT(1,2,3,4)");
        }

        [TestMethod]
        public void TestConcat()
        {
            test("\"1\"&\"2\"&\"3\"", "CONCATENATE(\"1\",\"2\",\"3\")");
        }

        private void cannotRefactor(string f)
        {
            Assert.IsFalse(testee.CanRefactor(f.Parse()), "Cannot refactor '{0}' but CanRefactor returned true", f);
        }

        private void test(string or, string target)
        {
            var orParsed = or.Parse();
            var targetParsed = target.Parse();

            Assert.IsTrue(testee.CanRefactor(orParsed), "Should be able to refactor '{0}' but returned false", or);

            var refactored = testee.Refactor(orParsed);

            Assert.AreEqual(Context.Empty.ProvideContext(targetParsed), Context.Empty.ProvideContext(refactored));
        }
    }
}
