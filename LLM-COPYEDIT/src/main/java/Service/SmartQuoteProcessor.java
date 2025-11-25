package Service;


import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;

public class SmartQuoteProcessor {

    public void process(String inputFilePath, String outputFilePath) throws Exception {
        // Load and parse the XML file
        File inputFile = new File(inputFilePath);
        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        dbFactory.setNamespaceAware(true);
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(inputFile);
        doc.getDocumentElement().normalize();

        // Get w:body node
        NodeList wbodyList = doc.getElementsByTagName("w:body");
        if (wbodyList.getLength() == 0) {
            throw new RuntimeException("No <w:body> found");
        }
        Node wbody = wbodyList.item(0);

        // Process every <w:t> node inside <w:body>
        NodeList tNodes = ((Element) wbody).getElementsByTagName("w:t");
        for (int i = 0; i < tNodes.getLength(); i++) {
            Node node = tNodes.item(i);
            if (node.getTextContent() != null) {
                node.setTextContent(smartQuotesExceptCant(node.getTextContent()));
            }
        }

        // Ensure the output directory exists
        File outputFile = new File(outputFilePath);
        File outputDir = outputFile.getParentFile();
        if (outputDir != null && !outputDir.exists()) {
            outputDir.mkdirs();
        }

        // Save result to specified output file
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        DOMSource source = new DOMSource(doc);
        StreamResult result = new StreamResult(outputFile);
        transformer.transform(source, result);

        System.out.println("File saved as: " + outputFilePath);
    }

    // Only replace quotes, preserve "can't"
    public static String smartQuotesExceptCant(String text) {
        if (text == null) return null;
        String placeholder = "__CANT_SMARTQUOTE_PLACEHOLDER__";
        // Protect all "can't" (case-insensitive)
        text = text.replaceAll("(?i)can't", placeholder);

        StringBuilder sb = new StringBuilder();
        boolean doubleOpen = true;
        boolean singleOpen = true;

        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch == '"') {
                sb.append(doubleOpen ? '“' : '”');
                doubleOpen = !doubleOpen;
            } else if (ch == '\'') {
                boolean prevIsLetterOrDigit = i > 0 && Character.isLetterOrDigit(text.charAt(i - 1));
                boolean nextIsLetterOrDigit = i + 1 < text.length() && Character.isLetterOrDigit(text.charAt(i + 1));
                
                if (prevIsLetterOrDigit && nextIsLetterOrDigit) {
                    // Apostrophe in contractions, like don't, we'll, etc.
                    sb.append('’'); // right single quote
                } else if (prevIsLetterOrDigit && !(nextIsLetterOrDigit)) {
                    // Apostrophe at end of word (possessives, like producers')
                    sb.append('’');
                } else if (!(prevIsLetterOrDigit) && nextIsLetterOrDigit) {
                    // Apostrophe at start of word (rare, as in 'tis)
                    sb.append('‘');
                } else {
                    // Paired quote for stand-alone uses
                    sb.append(singleOpen ? '‘' : '’');
                    singleOpen = !singleOpen;
                }
            } else {
                sb.append(ch);
            }
        }

        text = sb.toString();

        // Restore can't
        text = text.replace(placeholder, "can't");
        return text;
    }


    // Example usage:
    public static void main(String[] args) throws Exception {
        // Input your source file location and output file location:
        String inputFile = "C:\\Users\\Admin\\Downloads\\T_ECS1390156_CLN.xml";
        String outputFile = "D:\\CLN OUT\\output.xml";
        new SmartQuoteProcessor().process(inputFile, outputFile);
    }
}
