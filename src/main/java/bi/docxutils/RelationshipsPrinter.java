package bi.docxutils;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import static java.util.Comparator.comparing;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;

class RelationshipsPrinter
{
  public static void main(String[] args) throws FileNotFoundException, Docx4JException, IOException
  {
    if (args.length != 1 && args.length != 2)
      throw new RuntimeException("Expected 1 or 2 arguments: <docx-input-file> [output-file]");

    WordprocessingMLPackage docx = WordprocessingMLPackage.load(new File(args[0]));

    try (BufferedWriter bw = args.length >= 2
          ? new BufferedWriter(new FileWriter(args[1]))
          : new BufferedWriter(new OutputStreamWriter(System.out)))
    {
      writeRelationships(docx, bw);
    }
  }

  public static void writeRelationships(WordprocessingMLPackage docx, BufferedWriter bw) throws IOException
  {
    RelationshipsPart relsPart = docx.getMainDocumentPart().getRelationshipsPart(false);

    var rels = relsPart.getJaxbElement();

    for (var r : rels.getRelationship().stream().sorted(comparing(Relationship::getId)).toList())
    {
      bw.write(relDescr(r));
      bw.write("\n");
    }
  }

  private static String relDescr(Relationship r)
  {
    return new StringBuilder()
      .append("id: ").append(r.getId())
      .append(", type: ").append(r.getType())
      .append(", target: ").append(r.getTarget())
      .append(", target mode: ").append(r.getTargetMode())
      .toString();
  }
}
