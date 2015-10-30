using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using BumbleBee.Refactorings.Util;
using System.Windows.Forms;
using XLParser;
using Excel = NetOffice.ExcelApi;
using ExcelRaw = Microsoft.Office.Interop.Excel;
using Irony.Parsing;
//using XLParser;

namespace BumbleBee.Refactorings
{
    public class IntroduceCellName : RangeRefactoring
    {
        public override void Refactor(ExcelRaw.Range toName)
        {
            string subject = (toName.Count > 1) ? "Range" : "Cell";
            string name = existingName(toName);
            if (name == "")
            {
                name = Microsoft.VisualBasic.Interaction.InputBox($"{subject} name:", "Introduce Name");
            }
            else
            {
                MessageBox.Show($"Range has name {name} already, replacing all references to this range with {name}.");
            }
            
            // Empty string if user cancels, or doesn't fill anything in
            if (name != "")
            {
                Refactor(toName, name);
            }
        }

        /// <summary>
        /// Name a cell, and replace all references to it with the name
        /// </summary>
        /// <exception cref="AggregateException">If any cells could not be inlined, with as innerexceptions the individual errors.</exception>
        public void Refactor(ExcelRaw.Range toName, string name)
        {
            ParseTreeNode parse;
            try
            {
                parse = Helper.Parse(name);
            }
            catch (ArgumentException)
            {
                // Parse error
                throw new ArgumentException($"Name {name} is not a valid name for a named range");
            }
            parse = parse.SkipToRelevant();
            if (!parse.Is(GrammarNames.NamedRange))
            {
                throw new ArgumentException($"Name {name} is not a valid name for a named range, because Excel interpets it as a {parse.Type()}");
            }

            // Set the name
            var Scope = toName.Worksheet;
            var Names = Scope.Names;

            ExcelRaw.Name newName = Names.Add(name, toName);

            Marshal.ReleaseComObject(newName);
            Marshal.ReleaseComObject(Names);
            Marshal.ReleaseComObject(Scope);

            // Perform refactoring
            var ctx = toName.CreateContext();
            Inline(toName, ctx, ctx.Parse(name));
        }

        protected override RangeShape.Flags AppliesTo => RangeShape.Flags.NonEmpty;

        private static void Inline(ExcelRaw.Range toInline, Context toInlineCtx, ContextNode toInlineAST)
        {
            // Gather dependencies
            var dependencies = InlineFormula.GetAllDirectDependents(toInline);

            var toInlineAddress = Helper.ParseCtx(toInline.Address[false, false], toInlineCtx);

            var errors = new List<Exception>();
            foreach (ExcelRaw.Range dependent in dependencies)
            {
                try
                {
                    var dependentAST = Helper.ParseCtx(dependent);
                    if (dependentAST.Node == null)
                    {
                        throw new InvalidOperationException($"Could not parse formula of {dependent.SheetAndAddress()}");
                    }

                    var newFormula = dependentAST.Replace(toInlineAddress, toInlineAST);
                    try
                    {
                        dependent.Formula = "=" + newFormula.Print();
                    }
                    catch (COMException e)
                    {
                        throw new InvalidOperationException(
                            $"Refactoring produced invalid formula '={newFormula.Print()}' from original formula '{dependentAST.Print()}' for cell {dependent.SheetAndAddress()}",
                            e);
                    }
                }
                catch (Exception e)
                {
                    errors.Add(e);
                }
                finally
                {
                    Marshal.ReleaseComObject(dependent);
                }
            }

            if (errors.Count > 0)
            {
                throw new AggregateException(
                    $"Could not replace references with name in all dependents:\n{String.Join("\n", errors.Select(e => e.Message))}",
                    errors);
            }

        }

        
        private static string existingName(ExcelRaw.Range range)
        {
            try
            {
                ExcelRaw.Name eName = range.Name;
                string name = eName.Name;
                Marshal.ReleaseComObject(eName);

                // Remove sheet name
                var exclmarkIndex = name.IndexOf("!");
                if (exclmarkIndex >= 0)
                {
                    name = name.Substring(exclmarkIndex+1);
                }
                
                return name;
            }
            catch (COMException)
            {
                return "";
            }
        }
    }
}