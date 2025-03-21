package bi.docxutils;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.math.BigInteger;
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
import org.docx4j.wml.CTCnf;
import org.docx4j.wml.CTEm;
import org.docx4j.wml.CTFramePr;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.CTString;
import org.docx4j.wml.CTTblCellMar;
import org.docx4j.wml.CTTblLayoutType;
import org.docx4j.wml.CTTblOverlap;
import org.docx4j.wml.CTTblPPr;
import org.docx4j.wml.CTTblPrBase;
import org.docx4j.wml.CTTblPrBase.TblStyle;
import org.docx4j.wml.CTTextEffect;
import org.docx4j.wml.CTTextboxTightWrap;
import org.docx4j.wml.CTVerticalJc;
import org.docx4j.wml.Color;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Document;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.FldChar;
import org.docx4j.wml.Highlight;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.P;
import org.docx4j.wml.P.Hyperlink;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.DivId;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.OutlineLvl;
import org.docx4j.wml.PPrBase.PBdr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.PPrBase.TextAlignment;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPrAbstract;
import org.docx4j.wml.RStyle;
import org.docx4j.wml.STThemeColor;
import org.docx4j.wml.STVerticalJc;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Tabs;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblGrid;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcMar;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.TcPrInner.HMerge;
import org.docx4j.wml.TcPrInner.TcBorders;
import org.docx4j.wml.TcPrInner.VMerge;
import org.docx4j.wml.Text;
import org.docx4j.wml.TextDirection;

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
      case P p ->
        { if (p.getPPr() instanceof PPr ppr) sb.append(" ppr: { ").append(paragraphDescr(ppr)).append(" }"); }
      case Text t ->
        sb.append(" \"").append(t.getValue()).append("\"");
      case R r ->
        sb.append(runPropsDescr(r.getRPr()));
      case Tc tc ->
        sb.append(" ").append(tableCellPropsDescr(tc.getTcPr()));
      case Tbl t ->
        sb.append(" tblpr: { ").append(tablePropsDescr(t.getTblPr())).append(" }, tblgrid: ").append(t.getTblGrid() instanceof TblGrid g ? tableGridDescr(g) : "null");
      case CTBookmark b ->
        sb.append(" id: ").append(b.getId()).append(", name: ").append(b.getName());
      case Hyperlink h ->
        sb.append(" anchor: ").append(h.getAnchor());
      case Graphic g ->
        sb.append(" ").append(graphicDescr(g));
      case Inline i ->
        sb.append(" extent: ").append(sizeDescr(i.getExtent())).append(" graphic: ").append(graphicDescr(i.getGraphic()));
      case Anchor a ->
        sb.append(" ").append(anchorDescr(a));
      case FldChar f ->
        sb.append(" chartype: ").append(f.getFldCharType());
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

  private static String tableCellPropsDescr(@Nullable TcPr pr)
  {
    if (pr == null)
      return "";

    var props = new ArrayList<String>();

    if (pr.getCnfStyle() instanceof CTCnf ctcnf)
      props.add("cnfstyle: " + ctcnf.getVal());
    if (pr.getGridSpan() instanceof GridSpan gspan)
      props.add("gridspan: " + gspan.getVal());
    if (pr.getNoWrap() instanceof BooleanDefaultTrue b)
      props.add("nowrap: " + b.isVal());
    if (pr.getShd() instanceof CTShd shd)
      props.add("shd: {color: " + shd.getColor() + ", fill: " + shd.getFill() + ", ...}");
    if (pr.getTcBorders() instanceof TcBorders b)
      props.add("tcborders: { " +
        "top: {" + borderDescr(b.getTop()) + "}" +
        ", right: {" + borderDescr(b.getRight()) + "}" +
        ", bottom: {" + borderDescr(b.getBottom()) + "}" +
        ", left: {" + borderDescr(b.getLeft()) + "} " +
        "}"
      );
    if (pr.getTcFitText() instanceof BooleanDefaultTrue b)
      props.add("tcfittext: " + b.isVal());
    if (pr.getTcMar() instanceof TcMar b)
      props.add("tcmar: {" +
        "top: " + String.valueOf(b.getTop()) +
        ", right: " + String.valueOf(b.getRight()) +
        ", bottom: " + String.valueOf(b.getBottom()) +
        ", left: " + String.valueOf(b.getLeft()) +
        "}"
      );
    if (pr.getTcW() instanceof TblWidth w)
      props.add("tcw: { type: " + w.getType() + ", width: " + w.getW() + " }");
    if (pr.getVAlign() instanceof CTVerticalJc va && va.getVal() instanceof STVerticalJc v)
      props.add("valign: " + v.toString());
    if (pr.getVMerge() instanceof VMerge vm)
      props.add("vmerge: " + vm.getVal());
    if (pr.getHMerge() instanceof HMerge hm)
      props.add("hmerge: " + hm.getVal());

    return String.join(", ", props);
  }

  private static String runPropsDescr(@Nullable RPrAbstract pr)
  {
    if (pr == null)
      return "";

    var sb = new StringBuilder();

    if (pr.getB() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" bold");
    if (pr.getBCs() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" bold-complex-script");
    if (pr.getCaps() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" caps");
    if (pr.getCs() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" complex-script");
    if (pr.getDstrike() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" double-strike");
    if (pr.getEmboss() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" emboss");
    if (pr.getI() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" italics");
    if (pr.getICs() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" italics-complex-script");
    if (pr.getImprint() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" imprint");
    if (pr.getNoProof() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" no-proof");
    if (pr.getOMath() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" omath");
    if (pr.getOutline() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" outline");
    if (pr.getRtl() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" rtl");
    if (pr.getShadow() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" shadow");
    if (pr.getSmallCaps() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" smallcaps");
    if (pr.getSnapToGrid() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" snaptogrid");
    if (pr.getSpecVanish() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" specvanish");
    if (pr.getSmallCaps() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" smallcaps");
    if (pr.getStrike() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" strike");
    if (pr.getVanish() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" vanish");
    if (pr.getWebHidden() instanceof BooleanDefaultTrue b && b.isVal())
      sb.append(" webhidden");
    if (pr.getColor() instanceof Color c)
      sb.append(" color: ").append(c.getVal());
    if (pr.getHighlight() instanceof Highlight h)
      sb.append(" highlight: ").append(h.getHexVal());
    if (pr.getBdr() instanceof CTBorder b)
      sb.append(" border: ").append(b.toString());
    if (pr.getEm() instanceof CTEm e)
      sb.append(" em: ").append(e.toString());
    if (pr.getEffect() instanceof CTTextEffect e)
      sb.append(" texteffect: ").append(e.getVal().name());
    if (pr.getRFonts() instanceof RFonts f)
      sb.append(" fonts: { ascii: \"").append(f.getAscii()).append("\", cs: \"").append(f.getCs()).append("\", hansi: \"").append(f.getHAnsi()).append("\", eastasia: \"").append(f.getEastAsia()).append("\" }");;
    if (pr.getRStyle() instanceof RStyle s)
      sb.append(" style: ").append(s.getVal());
    if (pr.getSz() instanceof HpsMeasure s)
      sb.append(" size: ").append(s.getVal());
    if (pr.getSzCs() instanceof HpsMeasure s)
      sb.append(" complex-script-size: ").append(s.getVal());

    return sb.toString();
  }

  private static String paragraphDescr(@Nullable PPr pr)
  {
    if (pr == null)
      return "null";

    List<String> props = new ArrayList();

    if (pr.getAdjustRightInd() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("adjustrightind");
    if (pr.getAutoSpaceDE() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("autospacede");
    if (pr.getAutoSpaceDN() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("autospacedn");
    if (pr.getBidi() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("bidi");
    if (pr.getCnfStyle() instanceof CTCnf cnf)
      props.add("cnf: " + cnf.getVal());
    if (pr.getCollapsed() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("collapsed");
    if (pr.getContextualSpacing() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("contextualspacing");
    if (pr.getDivId() instanceof DivId d && d.getVal() instanceof BigInteger i)
      props.add("divid: " + i);
    if (pr.getFramePr() instanceof CTFramePr fpr)
      props.add("framepr: { h: " + fpr.getH() + ", hspace" + fpr.getHSpace() + ", ... }");
    if (pr.getInd() instanceof Ind i)
      props.add("ind: { left: " + i.getLeft() + ", ... }");
    if (pr.getJc() instanceof Jc jc)
      props.add("jc: " + jc.getVal());
    if (pr.getKeepLines() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("keeplines");
    if (pr.getKeepNext() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("keepnext");
    if (pr.getKinsoku() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("kinsoku");
    if (pr.getMirrorIndents() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("mirrorindents");
    if (pr.getNumPr() instanceof NumPr npr)
      props.add("numpr: { numid: " + npr.getNumId() + ", ilvl: " + npr.getIlvl() + " }");
    if (pr.getOutlineLvl() instanceof OutlineLvl l)
      props.add("outlinelvl: " + l.getVal());
    if (pr.getOverflowPunct() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("overflowpunct");
    if (pr.getPageBreakBefore() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("pagebreakbefore");
    if (pr.getPBdr() instanceof PBdr pbdr)
      props.add("pbdr: { " +
        "top: " + borderDescr(pbdr.getTop()) + ", " +
        "right: " + borderDescr(pbdr.getRight()) + ", " +
        "bottom: " + borderDescr(pbdr.getBottom()) + ", " +
        "left: " + borderDescr(pbdr.getLeft()) + ", " +
        "bar: " + borderDescr(pbdr.getBar()) + ", " +
        "between: " + borderDescr(pbdr.getBetween()) +
        " }"
      );
    if (pr.getPStyle() instanceof PStyle s)
      props.add("pstyle: " + s.getVal());
    if (pr.getRPr() instanceof ParaRPr prpr)
      props.add("rpr: { " + runPropsDescr(prpr) + " }");
    if (pr.getSectPr() instanceof SectPr s)
      props.add("sectpr: " + s);
    if (pr.getShd() instanceof CTShd shd)
      props.add("shd: { color: " + shd.getColor() + ", fill: " + shd.getFill() + ", ... }");
    if (pr.getSnapToGrid() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("snaptogrid");
    if (pr.getSpacing() instanceof Spacing s)
      props.add("spacing: { " +
        "line: " + s.getLine() + ", " +
        "linerule: " + s.getLineRule() + ", " +
        "before: " + s.getBefore() + ", " +
        "beforelines: " + s.getBeforeLines() + ", " +
        "after: " + s.getAfter() + ", " +
        "afterlines: " + s.getAfterLines() +
        " }"
      );
    if (pr.getSuppressAutoHyphens() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("suppressautohyphens");
    if (pr.getSuppressLineNumbers() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("suppresslinenumbers");
    if (pr.getSuppressOverlap() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("suppressoverlap");
    if (pr.getTabs() instanceof Tabs t)
      props.add("tabs: " + t.getTab());
    if (pr.getTextAlignment() instanceof TextAlignment ta)
      props.add("textalignment: " + ta.getVal());
    if (pr.getTextboxTightWrap() instanceof CTTextboxTightWrap ttw)
      props.add("textboxtightwrap: " + ttw.getVal());
    if (pr.getTextDirection() instanceof TextDirection td)
      props.add("textdirection: " + td.getVal());
    if (pr.getTopLinePunct() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("toplinepunct");
    if (pr.getWidowControl() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("windowcontrol");
    if (pr.getWordWrap() instanceof BooleanDefaultTrue b && b.isVal())
      props.add("wordwrap");

    return String.join(", ", props);
  }


  private static String borderDescr(@Nullable CTBorder b)
  {
    if (b == null)
      return "null";

    List<String> props = new ArrayList();

    if (b.isFrame())
      props.add("frame");
    if (b.isShadow())
      props.add("shadow");
    if (b.getColor() instanceof String c)
      props.add("color: " + c);
    if (b.getSpace() instanceof BigInteger s)
      props.add("space: " + s);
    if (b.getSz() instanceof BigInteger s)
      props.add("size: " + s);
    if (b.getThemeColor() instanceof STThemeColor c)
      props.add("themecolor: " + c.name());
    if (b.getThemeShade() instanceof String s)
      props.add("themeshade: " + s);
    if (b.getThemeTint() instanceof String t)
      props.add("themetint: " + t);

    return String.join(", ", props);
  }

  private static String tablePropsDescr(TblPr pr)
  {
    if (pr == null)
      return "null";

    List<String> props = new ArrayList();

    if (pr.getJc() instanceof Jc jc)
      props.add("jc: " + jc.getVal());
    if (pr.getShd() instanceof CTShd shd)
      props.add("shd: {color: " + shd.getColor() + ", fill: " + shd.getFill() + ", ...}");
    if (pr.getTblBorders() instanceof TblBorders b)
      props.add("tcborders: { " +
        "top: {" + borderDescr(b.getTop()) + "}" +
        ", right: {" + borderDescr(b.getRight()) + "}" +
        ", bottom: {" + borderDescr(b.getBottom()) + "}" +
        ", left: {" + borderDescr(b.getLeft()) + "} " +
        ", insideh: {" + borderDescr(b.getInsideH()) + "} " +
        ", insidev: {" + borderDescr(b.getInsideV()) + "} " +
        "}"
      );
    if (pr.getTblCaption() instanceof CTString s)
      props.add("tblcaption: \"" + s.getVal() + "\"");
    if (pr.getTblCellMar() instanceof CTTblCellMar m)
      props.add("tblcellmar: { " + cellMarDescr(m) + " }");
    if (pr.getTblCellSpacing() instanceof TblWidth w)
      props.add("tblcellspacing: " + w.getW());
    if (pr.getTblDescription() instanceof CTString s)
      props.add("tbldescription: " + s.getVal());
    if (pr.getTblInd() instanceof TblWidth w)
      props.add("tblind: " + w.getW());
    if (pr.getTblLayout() instanceof CTTblLayoutType tlt)
      props.add("tbllayout: " + tlt.getType());
    if (pr.getTblOverlap() instanceof CTTblOverlap to)
      props.add("tbloverlap: " + to.getVal());
    if (pr.getTblpPr() instanceof CTTblPPr ppr)
      props.add("tblppr: " + ppr);
    if (pr.getTblStyle() instanceof TblStyle s)
      props.add("tblstyle: " + s.getVal());
    if (pr.getTblStyleColBandSize() instanceof CTTblPrBase.TblStyleColBandSize s)
      props.add("tblstylecolbandsize: " + s.getVal());
    if (pr.getTblStyleRowBandSize() instanceof CTTblPrBase.TblStyleRowBandSize s)
      props.add("tblstylerowbandsize: " + s.getVal());
    if (pr.getTblW() instanceof TblWidth w)
      props.add("tblw: " + w.getW());

    return String.join(", ", props);

  }

  private static String cellMarDescr(@Nullable CTTblCellMar m)
  {
    if (m == null)
      return "null";

    List<String> props = new ArrayList();

    if (m.getTop() instanceof TblWidth w)
      props.add("top: " + w.getW());
    if (m.getRight() instanceof TblWidth w)
      props.add("right: " + w.getW());
    if (m.getBottom() instanceof TblWidth w)
      props.add("bottom: " + w.getW());
    if (m.getLeft() instanceof TblWidth w)
      props.add("left: " + w.getW());

    return String.join(", ", props);
  }

  private static String tableGridDescr(TblGrid g)
  {
    if (g.getGridCol() == null)
      return "";
    else
      return g.getGridCol().stream().map(gc -> gc.getW().toString()).toList().toString();
  }


  private static String sizeDescr(CTPositiveSize2D extent)
  {
    return extent.getCx()/EMUS_PER_CM+ "cm x " + extent.getCy()/EMUS_PER_CM + " cm";
  }

  private static final int EMUS_PER_CM = 360000;

}
