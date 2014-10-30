import java.io.*;
import org.antlr.runtime.*;
import org.antlr.runtime.debug.DebugEventSocketProxy;


public class __Test__ {

    public static void main(String args[]) throws Exception {
        formulaLexer lex = new formulaLexer(new ANTLRFileStream("C:\\Users\\Felienne\\infotron2\\infotron\\Infotron.FormulaParser\\output\\__Test___input.txt", "UTF8"));
        CommonTokenStream tokens = new CommonTokenStream(lex);

        formulaParser g = new formulaParser(tokens, 49100, null);
        try {
            g.ADD();
        } catch (RecognitionException e) {
            e.printStackTrace();
        }
    }
}