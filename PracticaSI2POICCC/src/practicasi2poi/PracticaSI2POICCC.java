package practicasi2poi;

import java.io.IOException;

import java.util.List;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

/**
 *
 * @author David Gonzalez Alvarez
 * @author Santiago Martinez Lopez
 * @version 1.0
 */
public class PracticaSI2POICCC {

    public static void main(String[] args) throws IOException, ParserConfigurationException, TransformerException {
        ExcelManager manager = new ExcelManager();
        List<String> lista = manager.obtenerDNIs();
        manager.comprobarDNIs(lista);
        manager.crearXml();
        //manager.comprobarCCC();
        //manager.generarIBAN();
        manager.comprobarEmail("");
        //manager.obtenerCCCs();
        //manager.imprimeCCCs();
        //manager.comprobarCCC();
        //manager.imprimeCCCsCorrectos();
        manager.generarIBAN();
    }
}
