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
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTEm;
import org.docx4j.wml.CTTextEffect;
import org.docx4j.wml.Color;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Document;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.FldChar;
import org.docx4j.wml.Highlight;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.P.Hyperlink;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.RStyle;
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

    // Append any interesting attributes not part of nested content here, to be displayed on initial line.
    switch (o)
    {
      case Text t ->
        sb.append(" \"").append(t.getValue()).append("\"");
      case R r ->
        sb.append(runPropsDescr(r.getRPr()));
      case CTBookmark b ->
        sb.append(" (id: ").append(b.getId()).append(", name: ").append(b.getName()).append(")");
      case Hyperlink h ->
        sb.append(" (anchor: ").append(h.getAnchor()).append(")");
      case Graphic g ->
        sb.append(" (").append(graphicDescr(g)).append(")");
      case Inline i ->
        sb.append(" (").append(sizeDescr(i.getExtent())).append(") Graphic ").append(graphicDescr(i.getGraphic()));
      case Anchor a ->
        sb.append(" ").append(anchorDescr(a));
      case FldChar f ->
        sb.append(" (chartype: ").append(f.getFldCharType()).append(")");
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

  private static String runPropsDescr(@Nullable RPr rPr)
  {
    if (rPr == null)
      return "";

    var sb = new StringBuilder();

    if (rPr.getB() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" bold");
    if (rPr.getBCs() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" bold-complex-script");
    if (rPr.getCaps() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" caps");
    if (rPr.getCs() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" complex-script");
    if (rPr.getDstrike() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" double-strike");
    if (rPr.getEmboss() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" emboss");
    if (rPr.getI() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" italics");
    if (rPr.getICs() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" italics-complex-script");
    if (rPr.getImprint() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" imprint");
    if (rPr.getNoProof() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" no-proof");
    if (rPr.getOMath() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" omath");
    if (rPr.getOutline() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" outline");
    if (rPr.getRtl() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" rtl");
    if (rPr.getShadow() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" shadow");
    if (rPr.getSmallCaps() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" smallcaps");
    if (rPr.getSnapToGrid() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" snaptogrid");
    if (rPr.getSpecVanish() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" specvanish");
    if (rPr.getSmallCaps() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" smallcaps");
    if (rPr.getStrike() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" strike");
    if (rPr.getVanish() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" vanish");
    if (rPr.getWebHidden() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" webhidden");
    if (rPr.getColor() instanceof Color c)
      sb.append(" color: ").append(c.getVal());
    if (rPr.getHighlight() instanceof Highlight h)
      sb.append(" highlight: ").append(h.getHexVal());
    if (rPr.getBdr() instanceof CTBorder b)
      sb.append(" border: ").append(b.toString());
    if (rPr.getEm() instanceof CTEm e)
      sb.append(" em: ").append(e.toString());
    if (rPr.getEffect() instanceof CTTextEffect e)
      sb.append(" texteffect: ").append(e.getVal().name());
    if (rPr.getRFonts() instanceof RFonts f)
      sb.append(" fonts: { ascii: \"").append(f.getAscii()).append("\", cs: \"").append(f.getCs()).append("\", hansi: \"").append(f.getHAnsi()).append("\", eastasia: \"").append(f.getEastAsia()).append("\" }");;
    if (rPr.getRStyle() instanceof RStyle s)
      sb.append(" style: ").append(s.getVal());
    if (rPr.getSz() instanceof HpsMeasure s)
      sb.append(" size: ").append(s.getVal());
    if (rPr.getSzCs() instanceof HpsMeasure s)
      sb.append(" complex-script-size: ").append(s.getVal());

    return sb.toString();
  }

  private static String sizeDescr(CTPositiveSize2D extent)
  {
    return extent.getCx()/EMUS_PER_CM+ "cm x " + extent.getCy()/EMUS_PER_CM + " cm";
  }

  private static final int EMUS_PER_CM = 360000;
}
