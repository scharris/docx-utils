package bi.docxutils;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Document;
import org.docx4j.wml.Text;

import static bi.util.Nullables.nonNull;

class DocxMainPartBodyStructurePrinter
{
  public static void main(String[] args) throws FileNotFoundException, Docx4JException, IOException
  {
    if (args.length != 1 && args.length != 2)
      throw new RuntimeException("Expected 1 or 2 arguments: <docx-input-file> [output-file]");

    WordprocessingMLPackage docx = WordprocessingMLPackage.load(new File(args[0]));

    try (BufferedWriter bw = args.length >= 2 ? new BufferedWriter(new FileWriter(args[1])) : new BufferedWriter(new OutputStreamWriter(System.out)))
    {
      writeStructure(docx, bw);
    }
  }

  public static void writeStructure(WordprocessingMLPackage docx, BufferedWriter bw) throws IOException
  {
    Body templateBody = ((Document)docx.getMainDocumentPart().getJaxbElement()).getBody();

    for (Object bodyContentItem : templateBody.getContent())
    {
      bw.write(contentString(bodyContentItem, ""));
      bw.write("\n");
    }
  }

  private static String contentString(Object o, String linesPrefix)
  {
    StringBuilder sb = new StringBuilder();
    Object unwrapped = nonNull(XmlUtils.unwrap(o));

    sb.append(linesPrefix).append(unwrapped.getClass().getSimpleName());
    if (unwrapped instanceof Text t)
      sb.append(" \"").append(t.getValue()).append("\"");

    sb.append("\n");

    if (unwrapped instanceof ContentAccessor contentAccessor)
    {
      boolean pastFirst = false;
      for (Object childContent : contentAccessor.getContent())
      {
        if (pastFirst) sb.append("\n");
        else pastFirst = true;
        sb.append(contentString(childContent, linesPrefix + "  "));
      }
    }

    return sb.toString();
  }
}
