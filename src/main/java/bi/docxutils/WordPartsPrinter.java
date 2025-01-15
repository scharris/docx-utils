package bi.docxutils;

import java.io.IOException;
import java.util.HashMap;

import jakarta.xml.bind.JAXBElement;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.contenttype.ContentTypeManager;
import org.docx4j.openpackaging.packages.Filetype;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.parts.DefaultXmlPart;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.FontTablePart;
import org.docx4j.openpackaging.parts.WordprocessingML.OleObjectBinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.VbaDataPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@SuppressWarnings("all")
public class WordPartsPrinter {

    private static Logger log = LoggerFactory.getLogger(WordPartsPrinter.class);

    /**
     * @param args
     */
    public static void main(String[] args) throws Exception {

        // Configuration options
        boolean printContentTypes = true;

        // Load the Package as an OpcPackage, since this works for docx, pptx, and xlsx
        OpcPackage opcPackage = OpcPackage.load(new java.io.File(args[0]), Filetype.ZippedPackage);

        handlePkg(opcPackage, printContentTypes);
    }

    public static void handlePkg(OpcPackage opcPackage, boolean printContentTypes) {

        if (printContentTypes) {
            printContentTypes(opcPackage);
        }

        // List the parts by walking the rels tree
        RelationshipsPart rp = opcPackage.getRelationshipsPart();
        StringBuilder sb = new StringBuilder();
        printInfo(rp.getPartName().getName(), null, rp, sb, "");
        traverseRelationships(opcPackage, rp, sb, "    ");

        System.out.println(sb.toString());
    }

    /**
     * It is often useful to see this [Content_Types].xml
     */
    public static void printContentTypes(org.docx4j.openpackaging.packages.OpcPackage p) {

        ContentTypeManager ctm = p.getContentTypeManager();
        System.out.println(ctm.toString());
    }

    public static void printInfo(String parentName, Relationship r, Part p, StringBuilder sb, String indent) {

        String relationshipType = "";
        if (p.getSourceRelationships().size() > 0) {
            relationshipType = p.getSourceRelationships().get(0).getType();
        }

        if (r == null) {
            sb.append("\n" + indent + "Part " + p.getPartName() + " [" + p.getClass().getName() + "] " + relationshipType);
        } else {
            sb.append("\n" + indent + parentName + "'s " + r.getId() + " is " + p.getPartName() + " [" + p.getClass().getName() + "] " + relationshipType);
        }

        if (p instanceof JaxbXmlPart) {
            Object o = ((JaxbXmlPart) p).getJaxbElement();
            if (o instanceof jakarta.xml.bind.JAXBElement) {
                sb.append(" containing JaxbElement:" + XmlUtils.JAXBElementDebug((JAXBElement) o));
            } else {
                sb.append(" containing:" + o.getClass().getName());
            }
        } else if (p instanceof DefaultXmlPart) {
            try {
                org.w3c.dom.Document doc = ((DefaultXmlPart) p).getDocument();
                try {
                    Object o = XmlUtils.unmarshal(doc);
                    if (o instanceof jakarta.xml.bind.JAXBElement) {
                        sb.append(" containing JaxbElement:" + XmlUtils.JAXBElementDebug((JAXBElement) o));
                    } else {
                        sb.append(" containing:" + o.getClass().getName());
                    }
                } catch (jakarta.xml.bind.UnmarshalException e) {
                    sb.append(" containing raw root element:" + doc.getDocumentElement().getLocalName());
                }
            } catch (Exception e) {
                throw new RuntimeException(e); // was: e.printStackTrace();
            }

        }

        if (p instanceof OleObjectBinaryPart) {

            try {
                ((OleObjectBinaryPart) p).viewFile(false);
            } catch (IOException e) {
                throw new RuntimeException(e); // was: e.printStackTrace();
            }
        }

        if (p instanceof VbaDataPart) {
            System.out.println(((VbaDataPart) p).getXML());
        }

        if (p instanceof FontTablePart) {
            ((FontTablePart) p).processEmbeddings();
        }
    }

    /**
     * This HashMap is intended to prevent loops.
     */
    public static HashMap<Part, Part> handled = new HashMap<Part, Part>();

    public static void traverseRelationships(org.docx4j.openpackaging.packages.OpcPackage wordMLPackage,
            RelationshipsPart rp,
            StringBuilder sb, String indent) {

        // TODO: order by rel id
        for (Relationship r : rp.getRelationships().getRelationship()) {

            log.info("\nFor Relationship Id=" + r.getId()
                    + " Source is " + rp.getSourceP().getPartName()
                    + ", Target is " + r.getTarget()
                    + " type " + r.getType() + "\n");

            if (r.getTargetMode() != null
                    && r.getTargetMode().equals("External")) {

                sb.append("\n" + indent + "external resource " + r.getTarget()
                        + " of type " + r.getType());
                continue;
            }

            Part part = rp.getPart(r);

            printInfo(rp.getSourceP().getPartName().getName(), r, part, sb, indent);
            if (handled.get(part) != null) {
                sb.append(" [additional reference] ");
                continue;
            }
            handled.put(part, part);
            if (part.getRelationshipsPart(false) == null) {
            } else {
                traverseRelationships(wordMLPackage, part.getRelationshipsPart(false), sb, indent + "    ");
            }
        }
    }
}
