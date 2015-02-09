module FSharpEngine.FSharpTransform

//-----------------------------------
let TakeNList (source:list<'a>)(n:int):list<'a> = source |> Seq.take n |> List.ofSeq

let RemainderList (source:list<'a>)(n:int):list<'a> = 
 match List.length(source) with
    | m -> source |> Seq.skip n |> Seq.take (m-n) |> List.ofSeq

let MapMerge (f:list<'a> -> list<'b>)(l:list<list<'a>>):list<'b> = List.concat (List.map f l)

let rec SubsTakePinNparts (source:list<'a>)(p:int)(n:int):list<list<list<'a>>> = 
    match List.length(source) with
    | 0 -> List.Empty
    | m -> 
        match p with
        | p when p > m-(n-1) -> List.empty
        | p -> List.map (List.append [TakeNList source p]) (AllSubLists (RemainderList source p) p (n-1)) 

and AllSubLists (source:list<'a>)(p:int)(n:int):list<list<list<'a>>> = 
    match p with
    | 0 -> List.empty
    | x -> 
        match n with
        |1 -> [[source]]
        |n when n = List.length (source) -> [List.map (fun x -> [x]) source]
        |n -> List.append ((SubsTakePinNparts source) p n) (AllSubLists source (p-1) n)

and AllSubTotal (source:list<'a>)(n:int):list<list<list<'a>>> = AllSubLists source (List.length(source)-(n-1)) n 


//----------------------------------

/// Variable for row or column
type Var = char
/// Variable cell address
type VarOp = Var * Var

// Concrete cell address
type Cell = int * int
// Dynamic/Variable cell address
type DCell =  VarOp * VarOp

/// Either a concrete or a dynamic cell
type SuperCell = 
/// Concrete cell address
|C of Cell
/// Dynamic cell address
|D of DCell

type Formula = 
    /// A constant/Literal
    | Constant of string
    /// Concrete or dynamic cell
    | S of SuperCell
    /// Concrete or dynamic cell range
    | Range of SuperCell * SuperCell
    /// Concrete function call with argument list
    | Function of string * list<Formula>
    /// Dynamic Range
    | DRange of Var
    /// Dynamic function argument
    | DArgument of Var
    /// List of function arguments (?)
    | ArgumentList of list<Formula>


type mapElement = 
 | Ints of list<int>
 | Formula of Formula

type maps = Map<char,mapElement> option
type TwoLists = list<Formula> * list<Formula>

type CellMap = bool * Map<char,list<int>>
type FormulaMap = bool * Map<char,Formula>

let makeConstant s = Constant s
let makeCell i j = C(i,j)
let makeSuperCell x = S x
let makeDCell i j = D(i,j)
let makeDRange r = DRange r
let makeDArgument c = DArgument c
let makeRange (x,y) = Range (x,y)
let makeFormula (s:string) (x:list<Formula>) = Function (s,x)

let rec IsDynamic (f:Formula) = 
    match f with
    | Constant c -> false
    | S(C (i,j)) -> false
    | Range (a, b) -> IsDynamic (S a) || IsDynamic (S b)
    | S(D (i,j)) -> true
    | DRange s -> true
    | DArgument c -> true
    // This seems wrong, wouldn't a formula be dynamic if any of the arguments were dynamic? Instead of all arguments?
    | Function (_, arguments) | ArgumentList(arguments) -> List.forall(IsDynamic) arguments

let HasMap f:bool = 
    match f with
    | Some (x) -> true 
    | None -> false

/// Returns the map from an option, or an empty map on none
let GetMap f:Map<Var,mapElement> = 
    match f with
    | Some (x) -> x 
    | None -> Map.empty
    
let JoinMaps (p:maps)(q:maps):maps = Some(Map(Seq.concat [ (Map.toSeq (GetMap p)) ; (Map.toSeq (GetMap q)) ]))

let CanJoinMaps (p:maps)(q:maps): bool = 
    match p with
    | None -> false
    | Some (x) -> 
        match q with
        |None -> false
        | Some (y) -> Map.forall(fun key value -> (y.TryFind(key) = None || x.TryFind(key) = y.TryFind(key))) x

let TryJoinMaps (p:maps)(q:maps): maps = if CanJoinMaps p q then (JoinMaps p q) else None

let rec CanBeAppliedonArguments (list1:list<Formula>) (list2:list<Formula>):maps= 
    List.fold2 (fun acc elem1 elem2 -> 
    
        (if not(CanBeAppliedonRoot(elem1, elem2)) then None else  (if (HasMap acc) then TryJoinMaps acc (CanBeAppliedon(elem1,elem2)) else (None))))
    
        (Some(Map.empty)) list1 list2
    // a fold2 function on two lists applies the function that is the first argument (fun acc elem1 elem2) on the elements of the two lists
    // acc is the intermediary result

and FromMapperSub (from: list<Formula>, source:list<list<Formula>>):maps = //list will always be of equal length, as source has been split in len(form) parts
    match (from, source) with 
    | (h1::t1, h2::t2) -> TryJoinMaps (TryMapArgumentList(h1,h2)) (FromMapperSub(t1,t2))
    | ((h::t), _) -> None
    | ([],[]) -> Some(Map.empty)

and FromMapper (from: list<Formula>, source:list<list<Formula>>):maps = //list will always be of equal length, as source has been split in len(form) parts
    match (from, source) with
    | (h1::t1, h2::t2) -> TryJoinMaps (TryMapArgumentList(h1,h2)) (FromMapper(t1,t2))
    | ((h::t), _) -> None
    | ([],[]) -> Some(Map.empty)

and AllDynamicMaps (from: list<Formula>, source: list<Formula>) = List.filter (fun elem -> HasMap(FromMapperSub (from, elem))) (AllSubFormulas source (List.length from))

and FirstDynamicMap (from: list<Formula>, source: list<Formula>) = 
    match (List.length(AllDynamicMaps (from, source))) with 
    | 0 -> List.empty
    | n -> List.nth (AllDynamicMaps (from, source)) 0 

and DynamicArgumentMapper (from: list<Formula>, source: list<Formula>):maps = if (List.length from >List.length source) then None else FromMapper (from, FirstDynamicMap(from,source))

and AllSubFormulas (source:list<Formula>)(n:int):list<list<list<Formula>>> = AllSubTotal source n 

and TryMapArgumentList (d:Formula, source: list<Formula>):maps = 
    match d with
    | DArgument c -> Some(Map.ofSeq [ (c, Formula(ArgumentList (source)))] )     
    | _ -> 
        match List.length(source) with
        | 1 -> CanBeAppliedon (d,(List.nth source 0))
        | n -> None

and CanBeAppliedon (from:Formula,source:Formula):maps = 
    match (from, source) with
    | (S(C _), S(C _)) -> Some(Map.empty)

    | (Constant x, Constant y) -> if x=y then Some(Map.empty) else None

    | (S(D (c,d)), S(C (i,j))) -> TryJoinMaps (Some(Map.ofSeq [ ((fst c), Ints([i - (int (snd c)-48)]))]))  (Some(Map.ofSeq [ ((fst d), Ints([j - (int (snd d)-48)]))])) //TODO: support more ops. For now: just support +, so deduct the second var of the operation
    | (S(C (i,j)), S(D (c,d))) -> CanBeAppliedon  (S(D (c,d)), S(C (i,j))) //swap and recall

    | (DRange(s), S(C(i,j))) -> Some(Map.ofSeq [ (s, Ints ([i;j])); ])
    | (S(C (i,j)), DRange(s)) -> CanBeAppliedon (DRange(s), S(C(i,j))) //swap and recall

    | (DRange s , Range (C(i,j), C(k,l))) ->  Some(Map.ofSeq [ (s, Ints( [i;j;k;l])); ])
    | (Range (C(i,j), C(k,l)), DRange(s)) -> CanBeAppliedon  (DRange(s), Range (C(i,j), C(k,l))) //swap and recall

    | (DArgument c, x) -> Some(Map.ofSeq [ (c, Formula (x)) ])
    | (x, DArgument c2) -> CanBeAppliedon( DArgument c2,  x) //swap and recall

    | ((Range (x1, y1)), (Range (x2, y2))) -> CanBeAppliedonArguments [S(x1);S(y1)] [S(x2);S(y2)]

    | (Function(s1, arguments1), Function(s2, arguments2)) -> if s1.Equals(s2) then (if arguments1.Length <> arguments2.Length then DynamicArgumentMapper(arguments1,arguments2) else  CanBeAppliedonArguments arguments1 arguments2 ) else None
     
    | _ -> None

and CanBeAppliedonMap (from:Formula,source:Formula):Map<Var,mapElement> = GetMap (CanBeAppliedon(from,source))
and CanBeAppliedonRoot (from:Formula,source:Formula):bool = HasMap (CanBeAppliedon(from,source) )

and CanBeAppliedonBool (from:Formula,source:Formula):bool = 
    if CanBeAppliedonRoot(from,source) then true  
    else
        (match source with 
        | Function (s, list) -> List.exists(fun x -> CanBeAppliedonBool(from, x)) list
        | _ -> false)

let FindVar (map:Map<char,mapElement>) i = //we only call this method when its type is <char, Ints I> still need to figure out how to convince te type checker of that :)
    match map.TryFind(i) with
    | Some x -> 
        match x with 
        | Ints I ->  List.head(I)
        | _ -> 0   
    | None -> (int i) - 49//not found, int so return the int

let FindRange (map:Map<char,mapElement>) i = 
    match map.TryFind(i) with
    | Some x -> 
        match x with
        | Ints I ->  
             match List.length(I) with
                | 2 -> S(C (List.nth I 0, List.nth I 1))
                | 4 -> Range (C(List.nth I 0, List.nth I 1), C(List.nth I 2, List.nth I 3))      
    | None -> S(C(0,219))

let FindFormula (map:Map<char,mapElement>) y = 
    match map.TryFind(y) with 
    | Some x -> 
        match x with
        | Formula F -> F        
    | None -> S(C(350,0))
    
let GetCellFromFormula (f:Formula):SuperCell =
    match f with
    | S(x) -> x
    | _ -> C(1,105)  


let rec MapFormula (map:Map<char,mapElement>) (t:Formula) : Formula=   
    match t with
    | S(D (i,j)) -> S(C (FindVar map (fst i) + (int (snd i)-48), FindVar map (fst j) + (int (snd j)-48)))
    | DRange r -> FindRange map r
    | Function (s, list) -> Function (s, List.map (MapFormula map) list)
    | Range (x,y) -> Range (GetCellFromFormula(MapFormula map (S(x))),GetCellFromFormula(MapFormula map (S(y))))
    | DArgument y -> FindFormula map y
    | _ -> t
//    | Formula f -> 
//    | _ -> null


/// <summary> Apply a transformation on a formula</summary>
/// <param name="to'">Formula transformation target</param>
/// <param name="from">Formula transformation origin</param>
/// <param name="source">Formula to apply transformation on</param>
/// <returns>The transformed formula if the transformation could be applied, the unaltered formula if it could not</returns>
let rec ApplyOn to' from source:Formula = 
    if source = from then
        to'
    else 
        //if the function matches exactly, then return to.
        if CanBeAppliedonRoot(from,source) then
            MapFormula (CanBeAppliedonMap (from,source)) to'
        else
            (match source with
            //try application in arguments
            | Function (s, list) -> Function (s, list |> List.map (ApplyOn to' from))
            | _ -> source)

type Formula with
    member this.ApplyOn to' from = this |> ApplyOn to' from

let rec Contains (search:Formula) (subject:Formula) : bool =
    if search = subject then
        true
    else
        match subject with
            | Function (_, arguments) | ArgumentList (arguments) -> Seq.exists (Contains search) arguments
            | _ -> false

type Formula with
    /// Check if the AST contains a certain subtree
    member this.Contains search = this |> Contains search

// You'd think this would be better done by defining `map f ast`
// but how do you decide whether to go deeper into the tree at a Function or apply f to the function?
let rec ReplaceSubTree search replace subject : Formula =
    // Found it, so replace
    if subject = search then
        replace
    else
        let doArgs arguments = arguments |> List.map (ReplaceSubTree search replace)
        match subject with
            // Look deeper into the AST
            | Function (s, arguments) -> Function(s, doArgs arguments)
            | ArgumentList (arguments) -> ArgumentList(doArgs arguments)
            // No match, do nothing
            | _ -> subject

type Formula with
    /// Replace every occurence of an expression in an AST with another expression
    member this.ReplaceSubTree search replace = this |> ReplaceSubTree search replace

/// Check if a cell is part of a range
let IsCellInRange cell range =
    match range with
        | Range(C(startC, startR), C(endC, endR)) -> 
            (match cell with
                | S(C(c,r)) -> c >= startC && c <= endC && r >= startR && r <= endR
                | _ -> invalidArg "cell" "cell must be a concrete cell")
        | _ -> invalidArg "range" "range must be a concrete range"

let rec RangesInFormula = function
    | Range (_,_) as r -> [r]
    | Function (_, arguments) | ArgumentList(arguments) -> arguments |> List.collect RangesInFormula
    | _ -> []

type Formula with
    member this.Ranges = this |> RangesInFormula

let ContainsCellInRanges cell formula = formula |> RangesInFormula |> List.exists (IsCellInRange cell)