using System;
using System.CodeDom;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Irony.Parsing;
using XLParser;
using NLog;
using NLog.LayoutRenderers;
using Infotron.Util;


namespace BumbleBee.Refactorings.Util
{
    /// <summary>
    /// A ParseTreeNode bundled with the context to do comparisons
    /// </summary>
    public class ContextNode
    {
        public Context Ctx { get; private set; }
        public ParseTreeNode Node { get; private set; }

        public ContextNode(Context ctx, ParseTreeNode node)
        {
            Ctx = ctx;
            Node = node;
        }

        public bool Contains(ContextNode search)
        {
            return Contains(Node, search.Node, Ctx, search.Ctx);
        }

        private static bool Contains(ParseTreeNode subject, ParseTreeNode search, Context csubject, Context csearch)
        {
            return Equals(subject, search, csubject, csearch)
                || AllowedChildPermutations(subject).Any(perm => perm.Any(child => Contains(child, search, csubject, csearch)));
        }


        /// <summary>
        /// Replace one subtree with another
        /// </summary>
        public ContextNode Replace(ContextNode search, ContextNode replace)
        {
            return new ContextNode(Ctx, Replace(Node, search.Node, replace.Node, Ctx, search.Ctx, replace.Ctx));
        }

        private static ParseTreeNode Replace(ParseTreeNode subject, ParseTreeNode search, ParseTreeNode replace, Context csub, Context csearch, Context crepl)
        {
            // Match, return the replacement
            if (Equals(subject, search, csub, csearch)) return MoveTo(replace, crepl, csub);

            // No match and no children to replace
            if (subject.ChildNodes.Count == 0) return subject;

            var newChilds = subject.ChildNodes.Select(pt => Replace(pt, search, replace, csub, csearch, crepl));
            var replacement = CustomParseTreeNode.From(subject);
            replacement.SetChildNodes(newChilds);
            return replacement;
        }

        /// <summary>
        /// Move a node to a new context.
        /// Provide the minimal qualification needed for formulas, so this might remove file/sheet prefixes if they are now superfluous
        /// </summary>
        /// <example>
        /// "A1" (sheet1) moveto (sheet2) will produce "sheet1!A1" (sheet2)
        /// "sheet2!A1" (sheet1) moveto (sheet2) will produce "A1" (sheet2)
        /// </example>
        public ContextNode MoveTo(Context newCtx)
        {
            return newCtx.Equals(Ctx) ? this : new ContextNode(newCtx, MoveTo(Node, Ctx, newCtx));
        }

        private static ParseTreeNode MoveTo(ParseTreeNode pt, Context old, Context _new)
        {
            if (pt.Is(GrammarNames.Reference))
            {
                return _new.QualifyMinimal(old.Qualify(pt));
            }
            return pt.ChildNodes.Count == 0 ? pt : CustomParseTreeNode.From(pt).SetChildNodes(pt.ChildNodes.Select(x => MoveTo(x, old, _new)));
        }

        public override string ToString()
        {
            return $"{Print()} in {Ctx}";
        }

        public string Print()
        {
            return Node.Print();
        }

        public override bool Equals(object o)
        {
            var other = o as ContextNode;
            if (other != null)
            {
                return Equals(Node, other.Node, Ctx, other.Ctx);
            }
            var pt = o as ParseTreeNode;
            return pt != null && Equals(Node, pt, Ctx, Ctx);
        }

        internal static bool Equals(ParseTreeNode p1, ParseTreeNode p2, Context c1, Context c2)
        {
            if (ReferenceEquals(p1, null)) return false;
            if (ReferenceEquals(p2, null)) return false;
            if (ReferenceEquals(c1, null)) return false;
            if (ReferenceEquals(c2, null)) return false;
            if (ReferenceEquals(p1, p2) && c1.Equals(c2)) return true;

            p1 = RemoveNonEqualityAffectingNodes(p1);
            p1 = c1.Qualify(p1);
            p2 = RemoveNonEqualityAffectingNodes(p2);
            p2 = c2.Qualify(p2);

            return
                // Compare term
                ((p1.Term == null && p2.Term == null) || (p1.Term != null && p2.Term != null && p1.Term.Name == p2.Term.Name))
                // Compare token
                && ((p1.Token == null && p2.Token == null) || (p1.Token != null && p2.Token != null && p1.Token.ValueString == p2.Token.ValueString))
                // Compare children
                && AllowedChildPermutations(p1).Any(childs1 => ChildrenEqual(childs1, p2.ChildNodes, c1, c2));
        }

        private static bool ChildrenEqual(List<ParseTreeNode> childs1, List<ParseTreeNode> childs2, Context c1, Context c2)
        {
            return // Check if they have the same number of children
                   childs1.Count == childs2.Count
                   // And if all of the children are equal
                   && childs1.Zip(childs2, (cn1, cn2) => Equals(cn1, cn2, c1, c2)).All(p => p);
        }

        private static readonly ISet<string> _communativeBinOps = new HashSet<string>()
        {
            "+",
            "*",
            GrammarNames.TokenIntersect,
            "-",
            "<>"
        };

        /// <summary>
        /// Get all the allowed permutations of child nodes.
        /// </summary>
        private static IEnumerable<List<ParseTreeNode>> AllowedChildPermutations(ParseTreeNode pt)
        {
            // Check if pt is a commutative binop operation
            if (ExcelFormulaParser.IsBinaryOperation(pt)
                && _communativeBinOps.Contains(ExcelFormulaParser.GetFunction(pt)))
            {
                var reversed = new List<ParseTreeNode>(pt.ChildNodes);
                reversed.Reverse();
                return new[] {pt.ChildNodes, reversed};
            }
            else
            {
                // Only original childnodes order allowed
                return Enumerable.Repeat(pt.ChildNodes, 1);
            }
        }

        public override int GetHashCode()
        {
            return GetHashCode(Ctx, Node);
        }

        internal static int GetHashCode(Context Ctx, ParseTreeNode pt)
        {
            pt = RemoveNonEqualityAffectingNodes(pt);
            pt = Ctx.Qualify(pt);
            int hash = 17;
            unchecked
            {
                if (pt.Term != null) hash = (hash*7) + pt.Term.Name.GetHashCode();
                if (pt.Token != null) hash = (hash*7) + pt.Token.ValueString.GetHashCode();
                hash = pt.ChildNodes.Aggregate(hash, (current, child) => (current*7) + GetHashCode(Ctx, child));
            }
            return hash;
        }

        private static ParseTreeNode RemoveNonEqualityAffectingNodes(ParseTreeNode pt)
        {
            return pt.SkipToRelevant(false);
        }

        /// <summary>
        /// Return all ranges used in a specific formula this cell is contained in
        /// </summary>
        public IEnumerable<ParseTreeNode> CellContainedInRanges(ContextNode formula)
        {
            var node = Node.SkipToRelevant(false);
            return CellContainedInRanges(Ctx.Qualify(node), formula);
        }

        private static IEnumerable<ParseTreeNode> CellContainedInRanges(ParseTreeNode fqcellref, ContextNode formula)
        {
            if (!fqcellref.Is(GrammarNames.Reference) || !fqcellref.ChildNodes[1].Is(GrammarNames.Cell))
            {
                throw new ArgumentException("Must be a reference to a single cell", nameof(fqcellref));
            }
            return CellContainedInRanges(fqcellref, formula.Node, formula.Ctx);
        }

        private static IEnumerable<ParseTreeNode> CellContainedInRanges(ParseTreeNode fqcellref, ParseTreeNode formula, Context CtxF)
        {
            var cell = new Location(fqcellref.ChildNodes[1].Print());
            // Select all  references and qualify them
            var references = formula.GetReferenceNodes().Select(CtxF.Qualify).ToList();
            
            // Check the different types of ranges
            var ranges = formula.AllNodes().Where(reference => reference.MatchFunction(":"));
            var rangesc = ranges.Where(range =>
                {
                    var args = range.GetFunctionArguments().Select(ExcelFormulaParser.Print).ToList();
                    var start = new Location(args[0]);
                    var end = new Location(args[1]);
                    return cell.Row >= start.Row && cell.Row <= end.Row
                           && cell.Column >= start.Column && cell.Column <= end.Column;
                });
            var vranges = references.Where(reference =>
                    reference.ChildNodes[0].Is(GrammarNames.Prefix)
                    && reference.ChildNodes[1].Is(GrammarNames.VerticalRange)
            );
            var vrangesc = vranges.Where(reference =>
            {
                var vrange = reference.ChildNodes[1];
                var pieces = vrange.Print().Replace("$", "").Split(':');
                return cell.Column >= AuxConverter.ColToInt(pieces[0])
                    && cell.Column <= AuxConverter.ColToInt(pieces[1]);
            });
            var hranges = references.Where(reference =>
                    reference.ChildNodes[0].Is(GrammarNames.Prefix)
                    && reference.ChildNodes[1].Is(GrammarNames.HorizontalRange)
            );
            var hrangesc = hranges.Where(reference =>
            {
                var hrange = reference.ChildNodes[1];
                var pieces = hrange.Print().Replace("$", "").Split(':');
                return cell.Row >= (int.Parse(pieces[0]) - 1) && cell.Row <= (int.Parse(pieces[1]) - 1);
            });
            var combined = new[] {rangesc, vrangesc, hrangesc}.SelectMany(x => x);
            return combined;
        }

        /// <summary>
        /// All (qualified) references used in this ContextNode
        /// </summary>
        public IEnumerable<ParseTreeNode> References => Node.GetReferenceNodes().Select(Ctx.Qualify);

        /// <summary>
        /// Return all named ranges referenced in this contextnode
        /// </summary>
        public IEnumerable<NamedRangeDef> NamedRanges
        {
            get
            {
                return References
                    .Where(reference => reference.ChildNodes[0].Is(GrammarNames.Prefix)
                                        && reference.ChildNodes[1].Is(GrammarNames.NamedRange)
                    ).Select(reference =>
                    {
                        var prefix = reference.ChildNodes[0].GetPrefixInfo();
                        return new NamedRangeDef(
                            prefix.FileName // file
                            , prefix.Sheet // sheet
                            , reference.ChildNodes[1].Print() // name
                            );
                    });
            }
        }
    }

    /// <summary>
    /// Provides context neccessary to compare nodes
    /// </summary>
    // In the future additional context might be neccessary for other operations
    public class Context
    {
        public ParserSheetReference DefinedIn { get; private set; }

        /// <summary>
        /// Named ranges defined (SheetRef, name).
        /// </summary>
        public IEnumerable<NamedRangeDef> NamedRanges { get; private set; }

        public Context(ParserSheetReference definedIn, ISet<NamedRangeDef> namedRanges = null)
        {
            DefinedIn = definedIn;
            NamedRanges = namedRanges ?? new HashSet<NamedRangeDef>();
        }

        /// <summary>
        /// Add a workbook and sheet reference to a prefix if they aren't provided yet
        /// </summary>
        /// <remarks>Qualify because (book,sheet,name) is a fully qualified name</remarks>
        public ParseTreeNode Qualify(ParseTreeNode reference)
        {
            // Check if this reference can be qualified
            if (!isPrefixableReference(reference)) return reference;

            var referenced = reference.ChildNodes.First(node => !node.Is(GrammarNames.Prefix));
            bool hasPrefix = reference.ChildNodes.Any(node => node.Is(GrammarNames.Prefix));
            var prefix = reference.FirstOrNewChild(CustomParseTreeNode.NonTerminal(GrammarNames.Prefix));

            PrefixInfo prefixinfo = null;
            if (hasPrefix)
            {
                prefixinfo = prefix.GetPrefixInfo();
            }

            var file = prefix.FirstOrNewChild(CustomParseTreeNode.NonTerminal(GrammarNames.File, GrammarNames.TokenEnclosedInBrackets, $"[{DefinedIn.FileName}]"));
            var sheet = prefix.FirstOrNewChild(CustomParseTreeNode.Terminal(GrammarNames.TokenSheet, DefinedIn.Worksheet));

            // Named ranges can be both workbook-level and sheet-level and need additional logic
            if (referenced.ChildNodes.First().Is(GrammarNames.NamedRange))
            {
                var name = referenced.ChildNodes.First().ChildNodes.First().Token.ValueString;

                // If a sheet was already provided, either the file was provided or will correctly be filled by definition above
                if (!(prefixinfo != null && prefixinfo.HasSheet))
                {
                    bool isDefinedOnSheetLevel = NamedRanges.Contains(new NamedRangeDef(DefinedIn, name));
                    // If a file was provided but no sheet, it's a workbook-level definition
                    // If a sheet-level name is not defined, either a workbook-level name is defined,
                    //   or it's not defined and we assume the Excel default of workbook-level
                    if ((prefixinfo != null && prefixinfo.HasFile) || !isDefinedOnSheetLevel)
                    {
                        sheet = CustomParseTreeNode.Terminal(GrammarNames.TokenSheet, "");
                    }
                }
            }

            prefix = prefix.SetChildNodes(file, sheet);

            return CustomParseTreeNode.From(reference).SetChildNodes(prefix, referenced);
        }

        private static bool isPrefixableReference(ParseTreeNode reference)
        {
            // No qualifying to do if it's not a reference
            if (!reference.Is(GrammarNames.Reference)) return false;
            // No qualifying to do if it's a Functioncall or dynamic data exchange
            var relevant = reference.SkipToRelevant();
            var child = relevant.ChildNodes.Count > 0 ? relevant.ChildNodes[0] : null;
            if ((child?.IsFunction()??false) || (child?.Is(GrammarNames.DynamicDataExchange)??false)) return false;
            return true;
        }

        /// <summary>
        /// Remove superfluous qualification from a node
        /// </summary>
        public ParseTreeNode QualifyMinimal(ParseTreeNode reference)
        {
            // Check if this reference can be qualified
            if (!isPrefixableReference(reference)) return reference;

            var referenced = reference.ChildNodes.First(node => !node.Is(GrammarNames.Prefix));
            var prefix = reference.FirstChild(GrammarNames.Prefix);
            // No prefix, it's already minimal
            if (prefix == null) return reference;

            var prefixinfo = prefix.GetPrefixInfo();

            var childs = new ParseTreeNodeList();
            if (prefixinfo.HasFileName && prefixinfo.FileName != DefinedIn.FileName)
            {
                childs.Add(CustomParseTreeNode.NonTerminal(GrammarNames.File, GrammarNames.TokenEnclosedInBrackets, $"[{prefixinfo.FileName}]"));
                childs.Add(CustomParseTreeNode.Terminal(GrammarNames.TokenSheet, prefixinfo.Sheet));
            } else if (prefixinfo.HasSheet && prefixinfo.Sheet != DefinedIn.WorksheetClean)
            {
                childs.Add(CustomParseTreeNode.Terminal(GrammarNames.TokenSheet, DefinedIn.Worksheet));
            } 

            if (childs.Count > 0)
            {
                prefix = CustomParseTreeNode.From(prefix).SetChildNodes(childs);
                return CustomParseTreeNode.From(reference).SetChildNodes(prefix, referenced);
            }
            else
            {
                return CustomParseTreeNode.From(reference).SetChildNodes(referenced);
            }
            
        }

        private class PtComparer : IEqualityComparer<ParseTreeNode>
        {
            private readonly Context Ctx;

            public PtComparer(Context Ctx)
            {
                this.Ctx = Ctx;
            }

            public bool Equals(ParseTreeNode x, ParseTreeNode y)
            {
                return ContextNode.Equals(x, y, Ctx, Ctx);
            }

            public int GetHashCode(ParseTreeNode obj)
            {
                return ContextNode.GetHashCode(Ctx, obj);
            }
        }

        /// <summary>
        /// Equality comparer for ParseTreeNodes in current context.
        /// </summary>
        public IEqualityComparer<ParseTreeNode> Comparer => new PtComparer(this);

        public static Context Empty { get; } = new Context(new ParserSheetReference("",""));

        public ContextNode Parse(string formula)
        {
            return new ContextNode(this, ExcelFormulaParser.Parse(formula));
        }

        public override bool Equals(object obj)
        {
            var other = obj as Context;
            if (ReferenceEquals(other, this)) return true;
            if (ReferenceEquals(other, null)) return false;
            return DefinedIn.Equals(other.DefinedIn) && NamedRanges.Equals(other.NamedRanges);
        }

        public override int GetHashCode()
        {
            int hash = 3;
            unchecked
            {
                hash = (hash*7) + DefinedIn.GetHashCode();
                hash = (hash*7) + NamedRanges.GetHashCode();
            }
            return hash;
        }

        public ContextNode ProvideContext(ParseTreeNode pt)
        {
            return new ContextNode(this, pt);
        }

        public override string ToString() =>  $"[{DefinedIn.FileName}]{DefinedIn.Worksheet}";
    }

    public class NamedRangeDef : Tuple<ParserSheetReference, string>
    {
        public ParserSheetReference DefinedIn => Item1;
        public string Workbook => DefinedIn.FileName;
        public string Worksheet => DefinedIn.WorksheetClean;
        public string Name => Item2;

        public bool IsSheetLevel => Worksheet != "!";

        public NamedRangeDef(ParserSheetReference loc, string name) : base(loc, name) {}
        public NamedRangeDef(string workbook, string name) : this(new ParserSheetReference(workbook,""), name) {}
        public NamedRangeDef(string workbook, string sheet, string name) : this(new ParserSheetReference(workbook, sheet), name){}
    }

    public static class PtExtensions
    {
        internal static ParseTreeNode FirstChild(this ParseTreeNode pt, string type)
        {
            return pt.ChildNodes.FirstOrDefault(x => x.Is(type));
        }

        /// <summary>
        /// Return a copy of the first child, or return the provided default
        /// </summary>
        internal static CustomParseTreeNode FirstOrNewChild(this ParseTreeNode pt, CustomParseTreeNode def)
        {
            return pt.ChildNodes
                .Where(x => x.Is(def.Term.Name))
                .Select(CustomParseTreeNode.From)
                .DefaultIfEmpty(def)
                .First();
        }
    }


    /// <summary>
    /// ParseTreeNode class has private childnode setter, so this is the ugly workaround
    /// </summary>
    internal class CustomParseTreeNode : ParseTreeNode
    {
        // This as ParseTreeNode to access hidden members
        private readonly ParseTreeNode pt;

        public new ParseTreeNodeList ChildNodes
        {
            get { return pt.ChildNodes; }
            set { SetChildNodes(value); }
        }

        public CustomParseTreeNode SetChildNodes(IEnumerable<ParseTreeNode> values)
        {
            pt.ChildNodes.Clear();
            pt.ChildNodes.AddRange(values);
            return this;
        }

        public CustomParseTreeNode SetChildNodes(params ParseTreeNode[] values)
        {
            return SetChildNodes((IEnumerable<ParseTreeNode>)values);
        }

        private CustomParseTreeNode(NonTerminal term, SourceSpan span) : base(term, span)
        {
            pt = this;
        }

        private CustomParseTreeNode(Token t) : base(t)
        {
            pt = this;
        }

        /// <summary>
        /// Create a CustomParseTreeNode from an existing ParseTreeNode
        /// </summary>
        public static CustomParseTreeNode From(ParseTreeNode pt)
        {
            // Check for nonterminal or terminal term
            var nonterm = pt.Term as NonTerminal;
            if (nonterm != null)
            {
                var ret = new CustomParseTreeNode(nonterm, pt.Span);
                ret.SetChildNodes(pt.ChildNodes);
                return ret;
            }
            else
            {
                return new CustomParseTreeNode(pt.Token);
            }
        }

        public static CustomParseTreeNode Terminal(string name, string value = "")
        {
            return new CustomParseTreeNode(new Token(new Terminal(name), new SourceLocation(0, 0, 0), value, value));
        }

        public static CustomParseTreeNode NonTerminal(string name)
        {
            return new CustomParseTreeNode(new NonTerminal(name), new SourceSpan(new SourceLocation(0, 0, 0), 0));
        }

        public static CustomParseTreeNode NonTerminal(string name, string terminal, string value)
        {
            return NonTerminal(name).SetChildNodes(Terminal(terminal, value));
        }

        public override string ToString()
        {
            return this.Print();
        }
    }
}