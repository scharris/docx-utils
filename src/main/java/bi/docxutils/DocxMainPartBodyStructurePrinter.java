package bi.docxutils;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;

import org.checkerframework.checker.nullness.qual.Nullable;
import org.docx4j.XmlUtils;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.picture.Pic;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.mce.AlternateContent;
import org.docx4j.mce.AlternateContent.Choice;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Document;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P.Hyperlink;
import org.docx4j.wml.Text;

import static bi.util.Nullables.nonNull;

class DocxMainPartBodyStructurePrinter
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

  private static String contentString(Object contentObject, String linesPrefix)
  {
    StringBuilder sb = new StringBuilder();
    Object o = nonNull(XmlUtils.unwrap(contentObject));

    sb.append(linesPrefix).append(o.getClass().getSimpleName());

    switch (o)
    {
      case Text t ->
        sb.append(" \"").append(t.getValue()).append("\"");
      case CTBookmark b ->
        sb.append(" (").append(b.getName()).append(")");
      case Hyperlink h ->
        sb.append(" (").append(h.getAnchor()).append(")");
      case Graphic g ->
        sb.append(" (").append(graphicDescr(g)).append(")");
      case Inline i ->
        sb.append(" (").append(sizeDescr(i.getExtent())).append(") Graphic ").append(graphicDescr(i.getGraphic()));
      case Anchor a ->
        sb.append(" ").append(anchorDescr(a));
      default -> {}
    }

    sb.append("\n");

    @Nullable List<Object> nestedContent =
      switch (o)
      {
        case ContentAccessor ca -> ca.getContent();
        case AlternateContent ac -> new ArrayList<>(ac.getChoice());
        case Drawing d -> d.getAnchorOrInline();
        case Choice c -> c.getAny();
        default -> null;
      };

    if (nestedContent != null)
    {
      boolean pastFirst = false;
      for (Object childContent : nestedContent)
      {
          if (pastFirst) sb.append("\n");
          else pastFirst = true;
          sb.append(contentString(childContent, linesPrefix + "  "));
      }
    }

    return sb.toString();
  }

  private static String graphicDescr(Graphic g)
  {
    return "GraphicData Pic " + picDescr(g.getGraphicData().getPic());
  }

  private static String anchorDescr(Anchor a)
  {
    return new StringBuilder()
      .append("id: ").append(HexFormat.of().formatHex(a.getAnchorId()))
      .append(",  hidden : ").append(a.isHidden())
      .append(",  Graphic ").append(graphicDescr(a.getGraphic()))
      .toString();
  }

  private static String picDescr(@Nullable Pic pic)
  {
    if (pic == null) return "(null)";

    StringBuilder sb = new StringBuilder();
    var blip = pic.getBlipFill().getBlip();
      sb.append("CTBlipFillProperties CTBlip [embed: ").append(blip.getEmbed()).append("]");
    return sb.toString();
  }

  private static String sizeDescr(CTPositiveSize2D extent)
  {
    return extent.getCx()/EMUS_PER_CM+ "cm x " + extent.getCy()/EMUS_PER_CM + " cm";
  }

  private static final int EMUS_PER_CM = 360000;
}
