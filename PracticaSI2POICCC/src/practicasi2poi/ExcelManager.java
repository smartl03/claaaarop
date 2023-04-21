package practicasi2poi;

//import static com.sun.org.apache.xalan.internal.xsltc.compiler.util.Type.String;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.HashMap;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellUtil;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

/**
 *
 * @author David Gonzalez Alvarez
 * @author Santiago Martinez Lopez
 * @version 1.0
 */
public class ExcelManager {

    public List listaTrabajadoresVacios = new ArrayList<>();
    public List listaTrabajadoresRepetidos = new ArrayList<>();
    public static List<Trabajador> listaTrabajadoresError = new ArrayList<Trabajador>();
    public List<Integer> mapaXML = new ArrayList<Integer>();
    private final List<String> listaDeCCCs;
    private final List<String> listaDeCCCsCorrecta;
    private List<String> listaPaises;
    //private final List<Integer> listaDeCCCsCorrectaEnteros;

    private static final HashMap<Integer, String> mapaDigitosControl = new HashMap<Integer, String>();
    //private static final HashMap<Integer, String> digitosCCC = new HashMap<Integer, String>();

    private static void rellenarHash() {
        mapaDigitosControl.put(0, "T");
        mapaDigitosControl.put(1, "R");
        mapaDigitosControl.put(2, "W");
        mapaDigitosControl.put(3, "A");
        mapaDigitosControl.put(4, "G");
        mapaDigitosControl.put(5, "M");
        mapaDigitosControl.put(6, "Y");
        mapaDigitosControl.put(7, "F");
        mapaDigitosControl.put(8, "P");
        mapaDigitosControl.put(9, "D");
        mapaDigitosControl.put(10, "X");
        mapaDigitosControl.put(11, "B");
        mapaDigitosControl.put(12, "N");
        mapaDigitosControl.put(13, "J");
        mapaDigitosControl.put(14, "Z");
        mapaDigitosControl.put(15, "S");
        mapaDigitosControl.put(16, "Q");
        mapaDigitosControl.put(17, "V");
        mapaDigitosControl.put(18, "H");
        mapaDigitosControl.put(19, "L");
        mapaDigitosControl.put(20, "C");
        mapaDigitosControl.put(21, "K");
        mapaDigitosControl.put(22, "E");
    }

    /*private static void elHashDeLosCCC() {
        digitosCCC.put(10, "A");
        digitosCCC.put(11, "B");
        digitosCCC.put(12, "C");
        digitosCCC.put(13, "D");
        digitosCCC.put(14, "E");
        digitosCCC.put(15, "F");
        digitosCCC.put(16, "G");
        digitosCCC.put(17, "H");
        digitosCCC.put(18, "I");
        digitosCCC.put(19, "J");
        digitosCCC.put(20, "K");
        digitosCCC.put(21, "L");
        digitosCCC.put(22, "M");
        digitosCCC.put(23, "N");
        digitosCCC.put(24, "O");
        digitosCCC.put(25, "P");
        digitosCCC.put(26, "Q");
        digitosCCC.put(27, "R");
        digitosCCC.put(28, "S");
        digitosCCC.put(29, "T");
        digitosCCC.put(30, "U");
        digitosCCC.put(31, "V");
        digitosCCC.put(32, "W");
        digitosCCC.put(33, "X");
        digitosCCC.put(34, "Y");
        digitosCCC.put(35, "Z");
    }*/

    public ExcelManager() {
        this.listaDeCCCs = new ArrayList<>();
        this.listaDeCCCsCorrecta = new ArrayList<>();
        rellenarHash();
        //elHashDeLosCCC();
    }

    public void crearTrabajadoresError() throws IOException {
        FileInputStream f = new FileInputStream(
                "./src/practicasi2poi/sistemasinformacionii.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(f)) {
            XSSFSheet sheet = wb.getSheetAt(0);

            for (int i = 0; i < mapaXML.size(); i++) {
                if (i != 0) {
                    Row row = CellUtil.getRow(mapaXML.get(i), sheet);
                    Cell nifnie = CellUtil.getCell(row, 11);
                    Cell nombre = CellUtil.getCell(row, 10);
                    Cell primerapellido = CellUtil.getCell(row, 8);
                    Cell segundoapellido = CellUtil.getCell(row, 9);
                    Cell empresa = CellUtil.getCell(row, 6);
                    Cell categoria = CellUtil.getCell(row, 7);
                    Trabajador t = new Trabajador(nifnie.toString(), nombre.toString(), primerapellido.toString(), segundoapellido.toString(), empresa.toString(), categoria.toString());
                    listaTrabajadoresError.add(t);
                }
            }
        }
    }

    public List<String> obtenerDNIs() throws IOException {
        FileInputStream f = new FileInputStream(
                "./src/practicasi2poi/sistemasinformacionii.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook(f)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            List<String> lista = new ArrayList<>();

            Row row;
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                while (cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    if (cell.getColumnIndex() == 11) {
                        lista.add(cell.toString());
                    }
                }
            }
            return lista;
        }
    }

    public int obtenerColumna(String dni) throws IOException {
        int resultado = 0;
        FileInputStream f = new FileInputStream(
                "./src/practicasi2poi/sistemasinformacionii.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook(f)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            Row row;

            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                while (cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    if (cell.getColumnIndex() == 11 && dni.equals(cell.toString())) {
                        resultado = cell.getRowIndex();
                    }
                }
            }
        } catch (IOException e) {
            System.out.println("Error: " + e.getMessage());
        }
        return resultado;
    }

    public void comprobarDNIs(List<String> lista) throws IOException {
        actualizarLetra();
        for (String s : lista) {
            if (dniCorrecto(s) == 0) {
                // Error no llega a la longitud, hay que meterlo en el xml
                mapaXML.add(obtenerColumna(s));
            } else if (dniCorrecto(s) == 2 || dniCorrecto(s) == 3) {
                if (repetido(s)) {
                    mapaXML.add(obtenerColumna(s));
                }
            }
        }
        crearTrabajadoresError();
    }

    /**
     *
     * @return
     * @throws FileNotFoundException
     * @throws IOException
     */
    private List<String> obtenerCCCs() throws FileNotFoundException, IOException {
        FileInputStream file = new FileInputStream("./src/practicasi2poi/SistemasInformacionII.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheetOfExcel = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheetOfExcel.iterator();
        Row row;
        while(rowIterator.hasNext()) {
                row = rowIterator.next();
                Iterator<Cell> cellIterator = row.iterator();
                Cell cell;
                while(cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    if(cell.getColumnIndex() == 1 && cell.getRowIndex() != 0) {
                        this.listaDeCCCs.add(cell.toString());
                    }
                }
        }
        return this.listaDeCCCs;
    }

    private boolean repetidoCCCs(String CCC) {
        boolean resultado = false;
        for(String s : this.listaDeCCCs) {
            if(s.equals(CCC)) {
                resultado = true;
                break;
            }
        }
        return resultado;
    }

    private boolean erroneoCCCs(String CCC) {
        boolean resultado = true;
        if(CCC.matches("[0-9]{20}")) {
            resultado = false;
        }
        return resultado;
    }

    public boolean comprobarCCC() throws FileNotFoundException, IOException{
        boolean resultado = false;
        FileInputStream file = new FileInputStream("./src/practicasi2poi/SistemasInformacionII.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        Row row;

        while(rowIterator.hasNext()) {
            row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            Cell cell;
            while(cellIterator.hasNext()) {
                cell = cellIterator.next();
                if((this.erroneoCCCs(cell.toString()) == false  && this.repetidoCCCs(cell.toString()) == false && cell.getColumnIndex() == 1 && cell.getRowIndex() != 0) || (cell.getColumnIndex() == 0 && cell.getRowIndex() != 0)) {
                        resultado = true;
                        this.listaPaises.add(cell.toString());
                        this.listaDeCCCsCorrecta.add(cell.toString());
                        //this.listaDeCCCsCorrecta.forEach((String x) -> System.out.println(+x+));
                }
            }
        }

        return resultado;
    }

    /**
     *
     * @return 
     * @throws java.io.IOException
     */
    /*public void imprimeCCCs() throws IOException {
        List<String> listaCCCs;
        listaCCCs = new ArrayList<>();
        listaCCCs = this.obtenerCCCs();
        listaCCCs.forEach((s) -> {
            System.out.println(s);
        });
    }*/

    /*public void imprimeCCCsCorrectos() throws IOException {
        List<String> listaCCCsCorrec;
        listaCCCsCorrec = new ArrayList<>();
        listaCCCsCorrec = comprobarCCC();
        int i = 0;
        while (i < listaCCCsCorrec.size()) {
            //JOptionPane.showMessageDialog(null, "Entraaaaaaandoooo");
            System.out.println(listaCCCsCorrec.get(i));
            i++;
        }
    }*/

    public String generarIBAN() throws IOException {
        List<BigDecimal> listaDeCCCsCorrectaEnteros = new ArrayList<BigDecimal>();
        String elIBAN = "";
        boolean resultado = this.comprobarCCC();
        int i = 0;
        if (resultado == true) {
            while (i < this.listaDeCCCsCorrecta.size()) {
                elIBAN = this.listaDeCCCsCorrecta.get(i) + "142800";
                this.listaDeCCCsCorrecta.set(i, elIBAN);
                BigDecimal num;
                num = new BigDecimal(this.listaDeCCCsCorrecta.get(i));
                listaDeCCCsCorrectaEnteros.add(num);
                i++;
            }  
        } else {
            while (i < this.listaDeCCCsCorrecta.size()) {
                this.listaDeCCCsCorrecta.remove(i);
                i++;
            }
        }
        listaDeCCCsCorrectaEnteros.forEach((s) -> {
            System.out.println(s);
        });
        return "";
    }

    public boolean comprobarEmail(String email){
        return false;
    }

    public String generarEmail(String nombre, String apellido1, String apellido2, String empresa){
        String emailResultado = "";
        String[] nombreSplit = nombre.split("");
        String[] apellido1Split = apellido1.split("");
        String[] apellido2Split = apellido2.split("");

        emailResultado += nombreSplit[0];
        emailResultado += apellido1Split[0];
        emailResultado += apellido2Split[0];
        emailResultado += empresa;
        emailResultado += ".com";
        
        return emailResultado;
    }

    public String generarEmail(String nombre, String apellido1, String empresa){
        String emailResultado = "";
        String[] nombreSplit = nombre.split("");
        String[] apellido1Split = nombre.split("");

        emailResultado += nombreSplit[0];
        emailResultado += apellido1Split[0] + "@";
        // * Numero de repeticion empezando en 00
        emailResultado += empresa;
        emailResultado += ".com";
        
        return emailResultado;
    }

    public int dniCorrecto(String dni) {
        int resultado = 0;
        if (dni.matches("[0-9]{8}[A-HJ-NP-TV-Z||a-hj-np-tv-z]")) {
            if (obtenerDigitoDeControl(dni) == "o") {
                resultado = 3;
            } else {
                resultado = 1;
            }
        } else if (dni.matches("[XYZxyz][0-9]{7}[A-HJ-NP-TV-Z||a-hj-np-tv-z]")) {
            if (obtenerDigitoDeControl(dni) == "o") {
                resultado = 2;
            } else {
                resultado = 1;
            }
        } else {
            resultado = 0;
        }
        return resultado;
    }

    // ! No hay que hacerlo como yo lo estoy haciendo, solo meter en el xml para siguiente practica
    public void eliminarDuplicado(String dni) throws IOException {
        int contador = 0;

        File o = new File("./src/practicasi2poi/resources/SistemasInformacionIITrasEjecucion.xlsx");
        FileInputStream f = new FileInputStream("./src/practicasi2poi/sistemasinformacionii.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook(f)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            Row row;

            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                while (cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    if (cell.getColumnIndex() == 11 && dni.equals(cell.toString()) && contador != 0) {
                        cell.setCellValue(" ");
                    } else {
                        contador++;
                    }
                }
            }

            o.createNewFile();
            wb.write(new FileOutputStream(o));
        } catch (IOException e) {
            System.out.println("Error: " + e.getMessage());
        }
        f.close();
    }

    public void actualizarLetra() throws FileNotFoundException, IOException {
        File o = new File(
                "./src/practicasi2poi/resources/SistemasInformacionIITrasEjecucion.xlsx");
        FileInputStream f = new FileInputStream(
                "./src/practicasi2poi/sistemasinformacionii.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook(f)) {
            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            Row row;
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                while (cellIterator.hasNext()) {
                    cell = cellIterator.next();
                    if (cell.getColumnIndex() == 11 && dniCorrecto(cell.toString()) != 0) {
                        cell.setCellValue(obtenerNuevoDNI(cell.toString()));
                    }
                }
            }
            o.createNewFile();
            wb.write(new FileOutputStream(o));
        } catch (IOException e) {
            System.out.print(e.getMessage());
        }
        f.close();
    }

    public static String obtenerDigitoDeControl(String dni) {
        String[] digitosSeparados = dni.split("");
        String numeroString = "";
        String letra = digitosSeparados[digitosSeparados.length - 1];
        int numero = 0;
        int numeroDelDigitoDeControl = 0;

        for (int i = 0; i < digitosSeparados.length; i++) {
            if (i == 0 && digitosSeparados[i].matches("[XYZ]")) {
                switch (digitosSeparados[i]) {
                    case "X":
                        digitosSeparados[i] = "0";
                        break;
                    case "Y":
                        digitosSeparados[i] = "1";
                        break;
                    case "Z":
                        digitosSeparados[i] = "2";
                        break;
                }
            }
            if (!digitosSeparados[i].matches("[A-Z]")) {
                numeroString += digitosSeparados[i];
            }
        }

        numero = Integer.parseInt(numeroString);

        numeroDelDigitoDeControl = numero % 23;

        if (letra.equals(mapaDigitosControl.get(numeroDelDigitoDeControl))) {
            return "o";
        } else {
            return mapaDigitosControl.get(numeroDelDigitoDeControl);
        }
    }

    public static String obtenerNuevoDNI(String dni) {
        String[] digitosSeparados = dni.split("");
        String resultado = "";

        if (dni.length() == 9) {
            String digitoDeControl = obtenerDigitoDeControl(dni);
            if (digitosSeparados[0].matches("[0-9]") && digitoDeControl != "o") {
                digitosSeparados[8] = obtenerDigitoDeControl(dni);
            } else if (digitosSeparados[0].matches("[XYZ]") && digitoDeControl != "o") {
                digitosSeparados[8] = obtenerDigitoDeControl(dni);
            }
            for (int i = 0; i < digitosSeparados.length; i++) {
                resultado += digitosSeparados[i];
            }
        }
        return resultado;
    }

    public boolean repetido(String dni) throws IOException {
        List<String> lista = obtenerDNIs();
        int contador = 0;
        boolean resultado = false;

        for (String s : lista) {
            if (s.equals(dni) && contador == 0) {
                contador++;
            } else if (s.equals(dni) && contador != 0) {
                contador++;
                resultado = true;
            }
        }
        return resultado;
    }

    public void crearXml() throws ParserConfigurationException, TransformerConfigurationException, TransformerException {

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();
            DOMImplementation implementacion = builder.getDOMImplementation();
            Document document = implementacion.createDocument(null, "Errores", null);

            Element trabajadores = document.createElement("Trabajadores");
            int contador = 0; // Se mira mañana

            for (Trabajador t : listaTrabajadoresError) {
                Element trabajador = document.createElement("Trabajador");

                Element nifnie = document.createElement("NIF_NIE");
                Text nifnietxt = document.createTextNode(t.getNifnie());
                nifnie.appendChild(nifnietxt);
                trabajador.appendChild(nifnie);

                Element nombre = document.createElement("Nombre");
                Text nombretxt = document.createTextNode(t.getNombre());
                nombre.appendChild(nombretxt);
                trabajador.appendChild(nombre);

                Element primerapellido = document.createElement("PrimerApellido");
                Text primerapellidotxt = document.createTextNode(t.getPrimerApellido());
                primerapellido.appendChild(primerapellidotxt);
                trabajador.appendChild(primerapellido);

                Element segundoapellido = document.createElement("SegundoApellido");
                Text segundoapellidotxt = document.createTextNode(t.getSegundoApellido());
                segundoapellido.appendChild(segundoapellidotxt);
                trabajador.appendChild(segundoapellido);

                Element empresa = document.createElement("Empresa");
                Text empresatxt = document.createTextNode(t.getEmpresa());
                empresa.appendChild(empresatxt);
                trabajador.appendChild(empresa);

                Element categoria = document.createElement("Categoria");
                Text categoriatxt = document.createTextNode(t.getCategoria());
                categoria.appendChild(categoriatxt);
                trabajador.appendChild(categoria);

                trabajadores.appendChild(trabajador);
            }

            document.getDocumentElement().appendChild(trabajadores);

            // GENERA XML
            Source source = new DOMSource(document);

            // DONDE SE GUARDARA
            Result result = new StreamResult(new File("./src/practicasi2poi/resources/Errores.xml"));
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.transform(source, result);

        } catch (ParserConfigurationException e) {
            System.out.println("Error: " + e.getMessage());
        }
    }
}
