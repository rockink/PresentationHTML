/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.rockink.presentationhtml;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.TreeMap;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPoint2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraphProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTransform2D;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;

/**
 *
 * @author nirmal
 */
public class PresentationConvert {

//    TreeMap<String, String> docMain;
    TreeMap<String, ShapeProperties> phDim;
    XMLSlideShow ppt;
    ZipInputStream zin;
    ZipEntry entry;
    private String path;
    double slideWidth, slideHeight;
    private XSLFSlide slide;

    public PresentationConvert(String path) throws IOException {

        this.path = path;
        ppt = new XMLSlideShow(this.getClass().getResourceAsStream("/Opening.pptx"));
        phDim = new TreeMap<>();

    }

    public PresentationConvert(InputStream path) throws IOException {

        ppt = new XMLSlideShow(path);
        zin = new ZipInputStream(new BufferedInputStream(path));
        phDim = new TreeMap<>();

    }

    public TreeMap<String, String>[] getRelation(int i) throws IOException {

        TreeMap<String, String> slideDirectoryPic = new TreeMap<>();
        TreeMap<String, String> slideRelPicDirectory = new TreeMap<>();
        TreeMap<String, String> ret[] = new TreeMap[2];

        XSLFSlide z = ppt.getSlides()[i];
        List<POIXMLDocumentPart> res = z.getRelations();

        for (POIXMLDocumentPart re : res) {

            PackagePart pckpart = re.getPackagePart();
            String content = pckpart.getContentType();

            if (content.contains("png") || content.contains("jpeg") || content.contains("jpg") || content.contains("rels")) {

                InputStream is = pckpart.getInputStream();
                byte[] bytezz = IOUtils.toByteArray(is);

                String hawa = StringUtils.newStringUtf8(Base64.encodeBase64(bytezz, false));
                slideDirectoryPic.put(pckpart.getPartName() + "", hawa);
                slideRelPicDirectory.put(re.getPackageRelationship().getId(), pckpart.getPartName() + "");
            }

        }

        ret[0] = slideDirectoryPic;
        ret[1] = slideRelPicDirectory;
        return ret;

    }

    public int getSlideCount() {

        return ppt.getSlides().length;

    }

    public String getSlide(int index) throws IOException {

        TreeMap<String, String>[] slideMain = getRelation(index);
        TreeMap<String, String> slideMainRelation = slideMain[0];
        TreeMap<String, String> slidePicDirectory = slideMain[1];
        TreeMap<String, ShapeProperties> slideDim = new TreeMap<>();

        String color = "";
        XSLFSlide slide = ppt.getSlides()[index];
        setSlide(slide);

//        XSLFSlideMaster masterSlide = slide.getSlideLayout().getSlideMaster();
        XSLFBackground masterSlideBackground = slide.getSlideLayout().getSlideMaster().getBackground();
//        XSLFBackground masterSlideBackground = masterSlide.getBackground();

        if (masterSlideBackground != null) {
            XmlObject slideXML = masterSlideBackground.getXmlObject();

            if (slide.getXmlObject().getCSld().isSetBg()) {

                color = getColorValueFromSolidFill(slide.getXmlObject().getCSld().getBg().getBgPr().getSolidFill().xmlText());

            } else {

                color = getColorValueFromSolidFill(slideXML.xmlText());

            }

            String style = style = "<style>"
                    + ".nir" + index
                    + "{background-color:'#" + color
                    + "';}"
                    + "</style>";

        }

        CTGroupShape sptree = slide.getXmlObject().getCSld().getSpTree();

        //the picture and text along with it!!
        List<CTShape> slideSp = sptree.getSpList();
        List<CTPicture> slidePic = sptree.getPicList();

        String htmlText = getFromSp(slideSp, slideDim);
        String htmlImage = getFromPic(slidePic, slideDim, slideMainRelation, slidePicDirectory);

        return "<section style='background-color:#" + color + ";'>" + htmlImage + htmlText + "</section>";
    }

    private ShapeProperties getShape(CTTransform2D frm) {

        double x, y, cx, cy;

        CTPoint2D cordinates = frm.getOff();
        x = cordinates.getX();
        y = cordinates.getY();

        x = getX(x);
        y = getY(y);

        CTPositiveSize2D cs = frm.getExt();
        cx = cs.getCx();
        cx = getX(cx);
        cy = cs.getCy();
        cy = getY(cy);

        return new ShapeProperties(x, y, cx, cy);

    }

    private double getX(double x) {

        double div = (double) x / (double) 121920;
        return div;

    }

    private double getY(double y) {

        double div = (double) y / (double) 68580;
        return div;

    }

    private String getFromSp(List<CTShape> slideSp, TreeMap<String, ShapeProperties> slideDim) {

        String sr = "";

        for (CTShape sp : slideSp) {

            ShapeProperties shape = null;
            CTShapeProperties properties = sp.getSpPr();

            //shape of overall slidePart
            if (properties.isSetXfrm()) {

                shape = getShape(properties.getXfrm());

            }

            //everything for a line!!
            sr += getEverythingFromSpLine(properties, shape);
            sr += getEverythingFromTxBody(sp, properties, shape, slideDim);

        }

        return sr;
    }

    //it creates html from list of picture it gets
    //returns html img src content for html
    //image is set as path -> content
    private String getFromPic(List<CTPicture> slidePic, TreeMap<String, ShapeProperties> slideDim,
            TreeMap<String, String> slideMainRelation, TreeMap<String, String> slidePicDirectory) {

        String bulkImage = "";
        for (CTPicture pic : slidePic) {

            //gets specific shape-------------------------------------------------------------------------------------!!!
            ShapeProperties shape = null;

            CTShapeProperties properties = pic.getSpPr();
            String imageId = pic.getBlipFill().getBlip().xgetEmbed().getStringValue();

            if (slideDim.get(imageId) == null) {

                if (properties.isSetXfrm()) {

                    shape = getShape(properties.getXfrm());
                    slideDim.put(imageId, shape);

                }

                //this happens if the shape is defined in layout! or further!
                if (!properties.isSetXfrm()) {

//                    XSLFSlideLayout slideLayout = getSlideLayout();

                    //get splist because you can actually find the dimension here!!
                    for (CTPicture slideLayoutSp : getSlideLayout().getXmlObject().getCSld().getSpTree().getPicList()) {

                        //slideDim puts value based on relID and its shape class
                        if (slideLayoutSp.getNvPicPr().getNvPr().isSetPh()) {
                            slideDim.put(pic.getBlipFill().getBlip().xgetEmbed().getStringValue(),
                                    getShape(slideLayoutSp.getSpPr().getXfrm()));
                        }

                    }

                    shape = slideDim.get(imageId);

                }

            } else {

                shape = slideDim.get(imageId);
            }

            //------------------------------------HAS SPECIFIC SHAPE-----------------------------------------------------
            String id = pic.getBlipFill().getBlip().xgetEmbed().getStringValue();

            String path = slidePicDirectory.get(id);

            //this seems unnecessary!!
//            if (path.contains(".png") || path.contains(".jpeg") || path.contains(".jpg")) {
            String style = " style = 'position:absolute;"
                    + "left:" + shape.getX()
                    + "%;"
                    + "max-width:" + shape.getCx()
                    + "%;"
                    + "top:" + shape.getY()
                    + "%;"
                    + "max-height:" + shape.getCy()
                    + "%;'";

            StringBuilder sbs = new StringBuilder();

            sbs.append("data:image/png;base64,")
                    .append(slideMainRelation.get(path)); //slidemainRelation stores image data in treemap!

            String image = sbs.toString();

            bulkImage += "<img src=" + image + " " + style + "/>";

//            }
        }

        return bulkImage;

    }

    private void setSlide(XSLFSlide slide) {

        this.slide = slide;
    }

    private XSLFSlideLayout getSlideLayout() {

        return slide.getSlideLayout();
    }

    //my debugger!!
    private void error(String s) {

        System.out.println("---------------------------------------------------------");
        System.out.println(s);
        System.out.println("---------------------------------------------------------");

    }

    private ShapeProperties setXfrm(XSLFSlide slide, CTShape sp, TreeMap<String, ShapeProperties> slideDim) {

//        CTShapeNonVisual nvsp = sp.getNvSpPr();
        String phXML = sp.getNvSpPr().getNvPr().getPh().xmlText();
//        CTApplicationNonVisualDrawingProps nvp = nvsp.getNvPr();
//        CTPlaceholder ph = nvp.getPh();
//        String phXML = ph.xmlText();

        ShapeProperties shape = slideDim.get(phXML);

        if (shape == null) {

            List<CTShape> layoutSpList = slide.getSlideLayout().getXmlObject().getCSld().getSpTree().getSpList();

            for (CTShape layoutSp : layoutSpList) {

                if (layoutSp.getNvSpPr().getNvPr().isSetPh()) {

                    //it is the layout correspondance!!
                    String phXml = layoutSp.getNvSpPr().getNvPr().getPh() + "";

                    //if layout of this specific isn't stored, store it!!
                    if (slideDim.get(phXml) == null) {

                        if (layoutSp.getSpPr().isSetXfrm()) {
                            slideDim.put(phXml, getShape(layoutSp.getSpPr().getXfrm()));
                        }

                    }

                }

            }

        }

        shape = slideDim.get(phXML);

        return shape;

    }

    private String getEverythingFromSpLine(CTShapeProperties properties, ShapeProperties shape) {

        String sr = "";
        if (properties.isSetLn()) {

            //gets the line in the form of html!!
            sr += getLine(properties.getLn(), shape);

        }

        return sr;
    }

    private String getLine(CTLineProperties ln, ShapeProperties shape) {

        String sr = "";

        if (ln.isSetSolidFill()) {

            CTSolidColorFillProperties sf = ln.getSolidFill();

            //color selector
            if (sf.isSetSrgbClr()) {

                String str = sf.xmlText();
                String[] splitStringArray = str.split("=| ");
                String color = getcolorValue(sf.xmlText());
                for (int i = 0; i < splitStringArray.length; i++) {

                    if (splitStringArray[i].equals("val")) {

                        color = splitStringArray[++i];
                        color = color.replace("\"", "");
                        break;

                    }

                }

                sr += "<hr style='position:absolute; left:" + shape.getX() + "%; max-width:" + shape.getCx() + "%;top:" + shape.getY() + "%;max-height:"
                        + shape.getCy() + "%; border-color:#" + color.replace("\"", "") + "'>";

            }

        }

        return sr;
    }

    //text and its property //from sp!
    private String getEverythingFromTxBody(CTShape sp, CTShapeProperties properties,
            ShapeProperties shape, TreeMap<String, ShapeProperties> slideDim) {

        String sr = "";

        //is it is textbody!!
        if (sp.isSetTxBody()) {

            properties = sp.getSpPr();
            shape = getShape(properties, sp, slideDim);

            if (shape != null) {
                sr += "<div style='position:absolute; left:"
                        + shape.getX() + "%; width:" + shape.getCx()
                        + "%;top:" + shape.getY()
                        + "%;height:"
                        + shape.getCy()
                        + "%'>";
            }

            List<CTTextParagraph> para = sp.getTxBody().getPList();

            //getlistproperty adds html tag, if it is ordered list, ul or something that is not on top of my head! 
            //but it accesses property, so there may be broader field to add in here!
            String runn = getListProperty(para);

            for (CTTextParagraph parag : para) {

                //sub function 
                if (parag.isSetPPr()) {
                    if (parag.getPPr().isSetBuAutoNum()) {

                        //make it a list!!
                        runn += "<li>";

                    }
                }

                //this actually creates div for every paragraph it encounters!!
                StringBuilder pos = new StringBuilder().append("<div>");

                List<CTRegularTextRun> run = parag.getRList(); //getRunList
                pos.append(getULDots(parag))
                        .append(getRunText(run))
                        .append("</div>");

                runn += pos;

            }

            if (para.get(0).isSetPPr()) {
                if (para.get(0).getPPr().isSetBuAutoNum()) {

                    runn += "</ol>";

                }
            }

            sr += runn + "</div><br>";
        }

        return sr;

    }

    private ShapeProperties getShape(CTShapeProperties properties, CTShape sp, TreeMap<String, ShapeProperties> slideDim) {

        ShapeProperties shape = null;
        if (properties.isSetXfrm()) {

            shape = getShape(properties.getXfrm());

        }

        //this is the correct one!!!
        if (!properties.isSetXfrm()) {

            shape = setXfrm(slide, sp, slideDim);

        }

        return shape;

    }

    private String getRunTestWithProperties(CTRegularTextRun ran) {

        String col = getSolidFillColor(ran);

        int size = ran.getRPr().getSz();
        String html = "<span style='font-size:" + size / 30 + "%;"
                + "color:#" + col + "'>" + ran.getT() + "</span>";
        return html;

    }

    private String getColorValueFromSolidFill(String xmlText) {

        String color = "";
        String[] splitStringArray = xmlText.split("=| |/"); //modified to support solidFill;
        for (int i = 0; i < splitStringArray.length; i++) {

            if (splitStringArray[i].equals("val")) {

                color = splitStringArray[++i];
                color = color.replace("\"", "");
                break;

            }

        }

        return color;
    }

    private String getSolidFillColor(CTRegularTextRun ran) {

        if (ran.isSetRPr()) {

            CTTextCharacterProperties runProper = ran.getRPr();
            if (runProper.isSetSolidFill()) {

                CTSolidColorFillProperties solidFill = runProper.getSolidFill();
                if (solidFill.isSetSrgbClr()) {

                    return getColorValueFromSolidFill(solidFill.xmlText());
                }

            }

        }

        return "";

    }

    private String getcolorValue(String xmlText) {

        String[] splitStringArray = xmlText.split("=| ");
        String color = "";
        for (int i = 0; i < splitStringArray.length; i++) {

            if (splitStringArray[i].equals("val")) {

                color = splitStringArray[++i];
                color = color.replace("\"", "");
                return color;

            }

        }

        return null;
    }

    private String getListProperty(List<CTTextParagraph> para) {

        String property = "";

        if (para.get(0).isSetPPr()) {
            CTTextParagraphProperties ppr;

            if ((ppr = para.get(0).getPPr()).isSetBuAutoNum()) {

                String an = ppr.getBuAutoNum().xgetType().getStringValue();

                if (an.equals("arabicParenR")) {
                    property += "<ol type = '1'>";
                }
                if (an.equals("romanUcPeriod")) {
                    property += "<ol type = 'I'>";
                }

            }
        }

        return property;
    }

    private String getULDots(CTTextParagraph parag) {

        String pos = "";

        if (parag.isSetPPr()) { //this is suspicious with the end!! //if property is set
            if (parag.getPPr().isSetBuChar()) {

                //it checks for line!!
                pos += "<span>" + parag.getPPr().getBuChar().getChar() + "</span>";

            }

        }//---------HERE!! DOWN HERE!!

        return pos;
    }

    private String getRunText(List<CTRegularTextRun> run) {

        String pos = "";
                        //this is the run list!!
        //this is ACTUALLY ALL THE TEXT THING COMMING FROM!! THIS IS IT!!
        for (CTRegularTextRun ran : run) {

            //to not go down and do stuff, if it isn't present. don't do it!!
            if (ran.getT().length() == 0) {
                continue;
            }

            if (ran.getRPr().isSetSz()) {

                pos += getRunTestWithProperties(ran);

            } else {

                String solCol = getSolidFillColor(ran);
                String parad = "<p class='a" + solCol + "'>" + ran.getT()
                        + "<br><style> .a" + solCol + "{color:#" + solCol + ";}</style>";
                pos += parad;

            }
        }

        return pos;
    }

}
