using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using FSharpEngine;
using Infotron.Parsing;
using Infotron.FSharpFormulaTransformation;

namespace TransformationTests
{
    [TestClass]
    public class FSharpASTReplacementTests
    {
        private readonly ExcelFormulaParser P = new ExcelFormulaParser();
        private readonly FSharpTransformationRule T = new FSharpTransformationRule();

        [TestMethod]
        public void ReplaceConstant()
        {
            TestReplace("1", "1", "2", "2");
        }

        [TestMethod]
        public void ReplaceCell()
        {
            TestReplace("D5", "D5", "E19", "E19");
        }

        [TestMethod]
        public void DontReplaceDifferentCell()
        {
            TestReplace("D6", "D5", "E19", "D6");
        }

        [TestMethod]
        public void ReplaceCellAbs()
        {
            TestReplace("$D$5", "D5", "E19", "E19");
        }

        [TestMethod]
        public void ReplaceRange()
        {
            TestReplace("A1:A5", "A1:A5", "D1:D5", "D1:D5");
        }

        [TestMethod]
        public void ReplaceFunction()
        {
            TestReplace("ABS(1)", "ABS(1)", "1", "1");
        }

        [TestMethod]
        public void ReplaceConstantInFunction()
        {
            TestReplace("1 + 1", "1", "2", "2 + 2");
        }

        [TestMethod]
        public void ReplaceCellInFunction()
        {
            TestReplace("D5 + 1", "D5", "E19", "E19 + 1");
        }

        [TestMethod]
        public void ReplaceNested()
        {
            TestReplace("1 + 2 * 3", "2*3", "6", "1 + 6");
        }

        [TestMethod]
        public void ReplaceBracketed()
        {
            TestReplace("1 + (2-3)", "2-3", "-1", "1 + (-1)");
        }

        private void TestReplace(string subject, string replace, string replacement, string expected)
        {
            var Fsub = T.CreateFSharpTree(P.ParseToTree(subject).Root);
            var Frep = T.CreateFSharpTree(P.ParseToTree(replace).Root);
            var Frepmnt = T.CreateFSharpTree(P.ParseToTree(replacement).Root);
            var Fexp = T.CreateFSharpTree(P.ParseToTree(expected).Root);

            var result = Fsub.ReplaceSubTree(Frep, Frepmnt);

            Assert.AreEqual(Fexp, result);
        }
    }
}
