using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Infotron.Parsing;
using Irony.Parsing;
using Infotron.FSharpFormulaTransformation;
using Microsoft.FSharp;
using Microsoft.FSharp.Collections;
using FSharpEngine;


namespace TransformationTests
{
    [TestClass]
    public class FSharpTransformationTests
    {

        [TestMethod]
        public void ConvertCell()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            string Cell = "A1";
            ParseTree Original = P.ParseToTree(Cell);

            FSharpTransformationRule T = new FSharpTransformationRule();
            FSharpTransform.Formula F = T.CreateFSharpTree(Original.Root);
            
            Assert.IsNotNull(F);
        }

        [TestMethod]
        public void ConvertRange()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            string Cell = "A1:B7";
            ParseTree Original = P.ParseToTree(Cell);

            FSharpTransformationRule T = new FSharpTransformationRule();
            FSharpTransform.Formula F = T.CreateFSharpTree(Original.Root);

            Assert.IsNotNull(F);
        }

        [TestMethod]
        public void ConvertFunction()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            string Cell = "SUM(A1:B7)";
            ParseTree Original = P.ParseToTree(Cell);

            FSharpTransformationRule T = new FSharpTransformationRule();
            FSharpTransform.Formula F = T.CreateFSharpTree(Original.Root);

            Assert.IsNotNull(F);
        }

        [TestMethod]
        public void ConvertSheetReference()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            string Cell = "Sheet!A1";
            ParseTree Original = P.ParseToTree(Cell);

            FSharpTransformationRule T = new FSharpTransformationRule();
            FSharpTransform.Formula F = T.CreateFSharpTree(Original.Root);

            Assert.IsNotNull(F);
        }


        [TestMethod]
        public void Can_Apply_Normal_Cell_Reference()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1")).Root;

            Assert.IsTrue(T.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Can_Apply_Calculation()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1+A2");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1+A2")).Root;
            T.to = (P.ParseToTree("SUM(A1:A2)")).Root;

            Assert.IsTrue(T.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Can_not_Apply_Calculation()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1*A2");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1+A2")).Root;
            T.to = (P.ParseToTree("SUM(A1:A2)")).Root;

            Assert.IsFalse(T.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        [Ignore]
        public void Can_Not_Apply_In_SubFormulas()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("(A1+A2)*A5");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1+A2")).Root;
            T.to = (P.ParseToTree("SUM(A1:A2)")).Root;

            Assert.IsFalse(T.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Can_Apply_In_SubFormulas()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1+A2)");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("SUM(A1+A2)")).Root;
            T.to = (P.ParseToTree("A1+A2")).Root;

            Assert.IsTrue(T.CanBeAppliedonBool(Original.Root));

            var map = (T.CanBeAppliedonMap(Original.Root));
        }


        [TestMethod]
        public void Can_Apply_Dynamic_Cell_Reference()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("{i,j}");

            var Result = T.CanBeAppliedonBool(Original.Root);
            Assert.IsTrue(Result);

            var Map = T.CanBeAppliedonMap(Original.Root);

            Assert.IsTrue(Map.ContainsKey('i'));
            Assert.IsTrue(Map.ContainsKey('j'));

            Assert.AreEqual(2, Map.Count);
        }

        [TestMethod]
        public void Can_Apply_Dynamic_Range_On_Cell_Reference()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1)");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM({r})");
            T.to = null;

            Assert.IsTrue(T.CanBeAppliedonBool(Original.Root));           
            var Map = T.CanBeAppliedonMap(Original.Root);

            Assert.IsTrue(Map.ContainsKey('r'));

            var ValueList = Map.TryFind('r').Value;
            var IntList = (FSharpTransform.mapElement.Ints)ValueList;

            Assert.IsTrue(IntList.Item.Contains(0));
            Assert.AreEqual(2,IntList.Item.Count());
        }

        [TestMethod]
        public void Can_Apply_Dynamic_Range_On_Range_Reference()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A2:B5)");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM({r})");
            T.to = null;

            Assert.IsTrue(T.CanBeAppliedonBool(Original.Root));

            var Map = T.CanBeAppliedonMap(Original.Root);

            Assert.IsTrue(Map.ContainsKey('r'));

            var ValueList = Map.TryFind('r').Value;

            var IntList = (FSharpTransform.mapElement.Ints)ValueList;

            Assert.IsTrue(IntList.Item.Contains(0));
            Assert.IsTrue(IntList.Item.Contains(1));
            Assert.IsTrue(IntList.Item.Contains(4));
        }

        [TestMethod]
        public void Can_Apply_Double_Dynamic_Reference()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A2:B5)+C5");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM({r})+{i,j}");
            T.to = null;

            Assert.IsTrue(T.CanBeAppliedonBool(Original.Root));

            var Map = T.CanBeAppliedonMap(Original.Root);

            Assert.IsTrue(Map.ContainsKey('r'));

            var ValueList = Map.TryFind('r').Value;            
            var IntList = (FSharpTransform.mapElement.Ints)ValueList;


            Assert.IsTrue(IntList.Item.Contains(0));
            Assert.IsTrue(IntList.Item.Contains(1));
            Assert.IsTrue(IntList.Item.Contains(4));

            Assert.IsTrue(Map.ContainsKey('i'));

            var ValueListi = Map.TryFind('i').Value;
            var IntListi = (FSharpTransform.mapElement.Ints)ValueListi;
            Assert.IsTrue(IntListi.Item.Contains(2));

            Assert.IsTrue(Map.ContainsKey('j'));
            var ValueListj = Map.TryFind('j').Value;
            var IntListj = (FSharpTransform.mapElement.Ints)ValueListj;
            Assert.IsTrue(IntListj.Item.Contains(4));
        }

        [TestMethod]
        public void Can_Not_Apply_Different_Dynamic_Cells()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A3");
            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("{i,i}");
            T.to = null;

            Assert.IsFalse(T.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Can_Not_Apply_Different_Dynamic_Range()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A2:C7)+SUM(A2:C6)");
            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM({r})+SUM({r})");
            T.to = null;

            Assert.IsFalse(T.CanBeAppliedonBool(Original.Root));
        }


        [TestMethod]
        public void Can_Apply_Dynamic_Cell_Reference_Multiple_Places()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1+A2");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("{i,j}+{i,j+1}");
            T.to = null;

            Assert.IsTrue(T.CanBeAppliedonBool(Original.Root));
        }


        [TestMethod]
        public void Can_not_Apply_Dynamic_Cell_Reference_Multiple_Places()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1+A2");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("{i,j}+{i+1,j}");
            T.to = null;

            Assert.IsFalse(T.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Dynamic_Argument_Can_Be_Text()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("\"5\"");

            FSharpTransformationRule S1 = new FSharpTransformationRule();
            S1.from = S1.ParseToTree("[c]");

            Assert.AreEqual(true, S1.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Dynamic_Argument_Can_Be_Cell()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1");

            FSharpTransformationRule S1 = new FSharpTransformationRule();
            S1.from = S1.ParseToTree("[c]");

            Assert.AreEqual(true, S1.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Dynamic_Argument_Can_Be_Formula()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1)");

            FSharpTransformationRule S1 = new FSharpTransformationRule();
            S1.from = S1.ParseToTree("[c]");

            Assert.AreEqual(true, S1.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Range_with_dymanice_cells_matches_range_with_cells()
        {
            FSharpTransformationRule R = new FSharpTransformationRule();
            R.from = R.ParseToTree("{i,j}: {i,j+1}");

            Assert.AreEqual(true, R.CanBeAppliedonBool("A1:A2"));


        }





        // Transformation Tests-------------

        [TestMethod]
        public void Super_Simple_Transform()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1")).Root;
            T.to = (P.ParseToTree("A2")).Root;

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);

            Assert.AreEqual("A2", T.Print(Result));
        }

        [TestMethod]
        public void Simple_Transform()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1+A2");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1+A2")).Root;
            T.to = (P.ParseToTree("SUM(A1:A2)")).Root;

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);

            Assert.AreEqual("SUM(A1:A2)", T.Print(Result));
        }

        [TestMethod]
        public void Transform_in_Arguments()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("(A1+A2)/3");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1+A2")).Root;
            T.to = (P.ParseToTree("SUM(A1:A2)")).Root;

            bool x = T.CanBeAppliedonBool(Original.Root);

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);

            Assert.AreEqual("(SUM(A1:A2))/3", T.Print(Result));
        }

        [TestMethod]
        public void Double_Transform()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("(A1+A2)*(A1+A2)");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = (P.ParseToTree("A1+A2")).Root;
            T.to = (P.ParseToTree("SUM(A1:A2)")).Root;

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);

            Assert.AreEqual("(SUM(A1:A2))*(SUM(A1:A2))", T.Print(Result));
        }


        [TestMethod]
        public void Transform_Dynamic_Cell_Reference()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("{i,j}");
            T.to = T.ParseToTree("{i,5}");

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);
            ParseTreeNode Expected = (P.ParseToTree("A5")).Root;

            Assert.AreEqual("A5", T.Print(Result));
        }


        [TestMethod]
        public void Transform_Dynamic_Cell_Calculation()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("A1");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("{i,j}");
            T.to = T.ParseToTree("{i,j+1}");

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);
            ParseTreeNode Expected = (P.ParseToTree("A2")).Root;

            Assert.AreEqual("A2", T.Print(Result));
        }

        //[TestMethod]
        //public void Dynamic_Transform_in_arguments()
        //{
        //    ExcelFormulaParser P = new ExcelFormulaParser();
        //    ParseTree Original = P.ParseToTree("A1+B1");

        //    TransformationRule T = new TransformationRule();
        //    T.from = T.ParseToTree("{i,j}+{i+1,j}");
        //    T.to = T.ParseToTree("SUM({i,j}:{i+1,j})");

        //    FSharpTransform.Formula Result = T.ApplyOn(Original.Root);
        //    ParseTreeNode Expected = (P.ParseToTree("SUM(A1:B1)")).Root;

        //    Assert.AreEqual(TransformationRule.Print(Expected), TransformationRule.Print(Result));
        //}


        [TestMethod]
        public void Dynamic_Range_Transform()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1:B1)");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM({r})");
            T.to = T.ParseToTree("SUM({r})+1");

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);
            Assert.AreEqual("SUM(A1:B1)+1", T.Print(Result));
        }

        [TestMethod]
        public void Dynamic_Range_Transform_on_Argument_list()
        {
            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM([c])");
            T.to = T.ParseToTree("SUM([c])+1");

            Assert.AreEqual(true, T.CanBeAppliedonBool("SUM(A1,B1)"));

            Assert.AreEqual("SUM(A1,B1)+1", T.ApplyOn("SUM(A1,B1)"));
        }

        [TestMethod]
        public void Dynamic_Range_Transform_With_Cell()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A7)");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM({r})");
            T.to = T.ParseToTree("SUM({r})+1");

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);

            Assert.AreEqual("SUM(A7)+1", T.Print(Result));
        }


        [TestMethod]
        public void Merge_Ranges()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1,A2)");

            FSharpTransformationRule T = new FSharpTransformationRule();
            T.from = T.ParseToTree("SUM({i,j}, {i,j+1})");
           
            T.to = T.ParseToTree("SUM({i,j}:{i,j+1})"); //hier heb je een gewone range met dt=ynamische cellen

            FSharpTransform.Formula Result = T.ApplyOn(Original.Root);
            ParseTreeNode Expected = (P.ParseToTree("SUM(A1,A2)")).Root;

            Assert.AreEqual("SUM(A1:A2)", T.Print(Result));

        }

        [TestMethod]
        public void Merge_Some_Ranges()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1,A2,A3)");

            FSharpTransformationRule S = new FSharpTransformationRule();
            S.from = S.ParseToTree("SUM({i,j}, {i,j+1}, {r})");
            S.to = S.ParseToTree("SUM({i,j}:{i,j+1}, {r})");

            FSharpTransform.Formula Result = S.ApplyOn(Original.Root);

            Assert.AreEqual("SUM(A1:A2,A3)", S.Print(Result));
        }


        [TestMethod]
        public void SUM_COUNT_AVERAGE()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1:A5)/COUNT(A1:A5)");

            FSharpTransformationRule S = new FSharpTransformationRule();
            S.from = S.ParseToTree("SUM({r})/COUNT({r})");
            S.to = S.ParseToTree("AVERAGE(A1:A5)");

            FSharpTransform.Formula Result = S.ApplyOn(Original.Root);

            Assert.AreEqual("AVERAGE(A1:A5)", S.Print(Result));
        }




        [TestMethod]
        public void Merge_Three_Ranges()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();

            FSharpTransformationRule R = new FSharpTransformationRule();
            R.from = R.ParseToTree("SUM({i,j}: {i,j+1}, {i,j+2})");
            R.to = R.ParseToTree("SUM({i,j}:{i,j+2})");
    
            Assert.AreEqual(true, R.CanBeAppliedonBool("SUM(A1:A2,A3)"));

            Assert.AreEqual("SUM(A1:A3)", R.ApplyOn("SUM(A1:A2,A3)"));
        }

        [TestMethod]
        public void Repeat_Merge()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("SUM(A1,A2,A3,A4)");

            FSharpTransformationRule S1 = new FSharpTransformationRule();
            S1.from = S1.ParseToTree("SUM({i,j}, {i,j+1})");
            S1.to = S1.ParseToTree("SUM({i,j}:{i,j+1})");

            Assert.AreEqual(false, S1.CanBeAppliedonBool(Original.Root));
        }

        [TestMethod]
        public void Fancy_Merge()
        {
            ExcelFormulaParser P = new ExcelFormulaParser();
            ParseTree Original = P.ParseToTree("(SUM(K3:K4,K5,K6,K7))/COUNT(K3:K7)");

            FSharpTransformationRule S1 = new FSharpTransformationRule();
            S1.from = S1.ParseToTree("SUM({x,y}: {i,j}, {i,j+1},[k])");
            S1.to = S1.ParseToTree("SUM({x,y}:{i,j+1},[k])");

            Assert.AreEqual(true, S1.CanBeAppliedonBool(Original.Root));
        }



        [TestMethod]
        public void Formula_With_Dynamic_Arguments()
        {
            FSharpTransformationRule T4 = new FSharpTransformationRule();
            T4.from = T4.ParseToTree("[c]+[d]");
            T4.to = T4.ParseToTree("SUM([c],[d])");

            Assert.AreEqual("SUM(A1,B1)", T4.ApplyOn("A1+B1"));
        }

        [TestMethod]
        public void Formula_With_Dynamic_Formula_Arguments()
        {
            FSharpTransformationRule T4 = new FSharpTransformationRule();
            T4.from = T4.ParseToTree("[c]+[d]");
            T4.to = T4.ParseToTree("SUM([c],[d])");

            Assert.AreEqual("SUM((A1*B1),(A2*B2))", T4.ApplyOn("(A1*B1)+(A2*B2)"));
        }


        [TestMethod]
        public void DynamicCell_Should_not_map_with_Function()
        {
            FSharpTransformationRule T4 = new FSharpTransformationRule();
            T4.from = T4.ParseToTree("SUM({i,j+1}, {i,j})");

            Assert.AreEqual(false, T4.CanBeAppliedonBool("SUM(A1,A8+A5)"));
        }

        [TestMethod]
        public void If_error()
        {

            FSharpTransformationRule S1 = new FSharpTransformationRule();
            S1.from = S1.ParseToTree("IF(ISERROR({r}),[c],{r})");
            S1.to = S1.ParseToTree("IFERROR({r},[c])");

            Assert.AreEqual("IFERROR(B2,\"Error\")", S1.ApplyOn("IF(ISERROR(B2),\"Error\",B2)"));
        }

        [TestMethod]
        public void If_error2()
        {
            FSharpTransformationRule S1 = new FSharpTransformationRule();
            S1.from = S1.ParseToTree("IF(ISERROR([d]),[c],[d])");
            S1.to = S1.ParseToTree("IFERROR([d],[c])");

            Assert.AreEqual("IFERROR(A1+A2+B1,\"Error\")", S1.ApplyOn("IF(ISERROR(A1+A2+B1),\"Error\",A1+A2+B1)"));
        }


        [TestMethod]
        public void SUM_in_argument_should_not_match_with_cell()
        {
            FSharpTransformationRule T6 = new FSharpTransformationRule();
            T6.from = T6.ParseToTree("SUM([c])+[e]");
            Assert.AreEqual(false, T6.CanBeAppliedonBool("A2+B1"));
        }


        [TestMethod]
        public void SUM_in_argument_should_match()
        {
            FSharpTransformationRule T6 = new FSharpTransformationRule();
            T6.from = T6.ParseToTree("[c]+[d]");

            Assert.AreEqual(true, T6.CanBeAppliedonBool("SUM(A1,A8+A5)"));
        }


        [TestMethod]
        public void SUM_in_argument_should_match2()
        {
            FSharpTransformationRule T6 = new FSharpTransformationRule();
            T6.from = T6.ParseToTree("SUM([c],SUM([d]))");
            T6.to = T6.ParseToTree("SUM([c], [d])");

            Assert.AreEqual(true, T6.CanBeAppliedonBool("SUM(A1,A8,SUM(A5))"));

            Assert.AreEqual("SUM(A1,A8,A5)", T6.ApplyOn("SUM(A1,A8,SUM(A5))"));
        }


        [TestMethod]
        public void SUM_in_argument_should_match3()
        {
            FSharpTransformationRule T6 = new FSharpTransformationRule();
            T6.from = T6.ParseToTree("SUM([c],SUM([d]))");
            T6.to = T6.ParseToTree("SUM([c], [d])");

            Assert.AreEqual(true, T6.CanBeAppliedonBool("SUM(A1,A8,A9,SUM(A5))"));

            Assert.AreEqual("SUM(A1,A8,A9,A5)", T6.ApplyOn("SUM(A1,A8,A9,SUM(A5))"));
        }



        [TestMethod]
        public void SUM_in_argument_should_match4()
        {
            FSharpTransformationRule T6 = new FSharpTransformationRule();
            T6.from = T6.ParseToTree("SUM([c],SUM([d]), [e])");
            T6.to = T6.ParseToTree("SUM([c], [d], [e])");

            Assert.AreEqual(true, T6.CanBeAppliedonBool("SUM(A1,SUM(A5), A4,A7)"));

            Assert.AreEqual("SUM(A1,A5,A4,A7)", T6.ApplyOn("SUM(A1,SUM(A5),A4,A7)"));
        }

        [TestMethod]
        public void SUM_in_argument_should_match5()
        {   

            FSharpTransformationRule S = new FSharpTransformationRule();
            S.from = S.ParseToTree("SUM({i,j+1}, [c], {i,j}, [d])");
            S.to = S.ParseToTree("SUM({i,j},[c],{i,j+1},[d])");

            Assert.AreEqual(false, S.CanBeAppliedonBool("SUM(A1,SUM(A8,SUM(A5,SUM(A3,A2))))"));
        }

        [TestMethod]
        public void SUM_in_argument_should_match6()
        {
            FSharpTransformationRule T6 = new FSharpTransformationRule();
            T6.from = T6.ParseToTree("SUM({i,j}:{i+2,j})/3");

            Assert.AreEqual(true, T6.CanBeAppliedonBool("SUM(B3:D3)/3"));
        }



        [TestMethod]
        public void Transforming_strings()
        {
            FSharpTransformationRule T6 = new FSharpTransformationRule();
            T6.from = T6.ParseToTree("[c]/[d]");
            T6.to = T6.ParseToTree("IF([d]<>0,[c]/[d],\"Error\")");

            Assert.AreEqual(true, T6.CanBeAppliedonBool("A1/A2"));

            string g = T6.ApplyOn("A1/A2");

            Assert.AreEqual("IF(A2<>0,A1/A2,\"Error\")", g);
        }




    }
}
