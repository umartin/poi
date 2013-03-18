/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.xssf.model;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBackgroundFillStyleList;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBevel;
import org.openxmlformats.schemas.drawingml.x2006.main.CTColorScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.ThemeDocument;
import org.openxmlformats.schemas.drawingml.x2006.main.CTColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTEffectList;
import org.openxmlformats.schemas.drawingml.x2006.main.CTEffectStyleItem;
import org.openxmlformats.schemas.drawingml.x2006.main.CTEffectStyleList;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFillStyleList;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGradientFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGradientStop;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineStyleList;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOuterShadowEffect;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSRgbColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTScene3D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSchemeColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTStyleMatrix;
import org.openxmlformats.schemas.drawingml.x2006.main.STCompoundLine;
import org.openxmlformats.schemas.drawingml.x2006.main.STLightRigDirection;
import org.openxmlformats.schemas.drawingml.x2006.main.STLightRigType;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineCap;
import org.openxmlformats.schemas.drawingml.x2006.main.STPathShadeType;
import org.openxmlformats.schemas.drawingml.x2006.main.STPenAlignment;
import org.openxmlformats.schemas.drawingml.x2006.main.STPresetCameraType;
import org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal;
import org.openxmlformats.schemas.drawingml.x2006.main.STSchemeColorVal;

/**
 * Class that represents theme of XLSX document. The theme includes specific
 * colors and fonts.
 * 
 * @author Petr Udalau(Petr.Udalau at exigenservices.com) - theme colors
 */
public class ThemesTable extends POIXMLDocumentPart {
    private ThemeDocument theme;

    public ThemesTable(PackagePart part, PackageRelationship rel) throws IOException {
        super(part, rel);
        
        try {
           theme = ThemeDocument.Factory.parse(part.getInputStream());
        } catch(XmlException e) {
           throw new IOException(e.getLocalizedMessage());
        }
    }

	public ThemesTable() {
		this(ThemeDocument.Factory.newInstance());
	}

    public ThemesTable(ThemeDocument theme) {
        this.theme = theme;
		initialize();
    }

    public XSSFColor getThemeColor(int idx) {
        CTColorScheme colorScheme = theme.getTheme().getThemeElements().getClrScheme();
        CTColor ctColor = null;
        int cnt = 0;
        for (XmlObject obj : colorScheme.selectPath("./*")) {
            if (obj instanceof org.openxmlformats.schemas.drawingml.x2006.main.CTColor) {
                if (cnt == idx) {
                    ctColor = (org.openxmlformats.schemas.drawingml.x2006.main.CTColor) obj;
                    
                    byte[] rgb = null;
                    if (ctColor.getSrgbClr() != null) {
                       // Colour is a regular one 
                       rgb = ctColor.getSrgbClr().getVal();
                    } else if (ctColor.getSysClr() != null) {
                       // Colour is a tint of white or black
                       rgb = ctColor.getSysClr().getLastClr();
                    }

                    return new XSSFColor(rgb);
                }
                cnt++;
            }
        }
        return null;
    }
    
    /**
     * If the colour is based on a theme, then inherit 
     *  information (currently just colours) from it as
     *  required.
     */
    public void inheritFromThemeAsRequired(XSSFColor color) {
       if(color == null) {
          // Nothing for us to do
          return;
       }
       if(! color.getCTColor().isSetTheme()) {
          // No theme set, nothing to do
          return;
       }

       // Get the theme colour
       XSSFColor themeColor = getThemeColor(color.getTheme());
       // Set the raw colour, not the adjusted one
       // Do a raw set, no adjusting at the XSSFColor layer either
       color.getCTColor().setRgb(themeColor.getCTColor().getRgb());

       // All done
    }

	private void initialize() {
		theme.addNewTheme();
		theme.getTheme().addNewThemeElements();

		// Set default color scheme.
		CTColorScheme clrScheme = theme.getTheme().getThemeElements().addNewClrScheme();
		setCTColor(clrScheme.addNewDk1(), 0, 0, 0);
		setCTColor(clrScheme.addNewLt1(), 0xff, 0xff, 0xff);
		setCTColor(clrScheme.addNewDk2(), 0x1f, 0x49, 0x7d);
		setCTColor(clrScheme.addNewLt2(), 0xee, 0xec, 0xe1);
		setCTColor(clrScheme.addNewAccent1(), 0x4f, 0x81, 0xbd);
		setCTColor(clrScheme.addNewAccent2(), 0xc0, 0x50, 0x4d);
		setCTColor(clrScheme.addNewAccent3(), 0x9b, 0xbb, 0x59);
		setCTColor(clrScheme.addNewAccent4(), 0x80, 0x64, 0xa2);
		setCTColor(clrScheme.addNewAccent5(), 0x4b, 0xac, 0xc6);
		setCTColor(clrScheme.addNewAccent6(), 0xf7, 0x96, 0x46);
		setCTColor(clrScheme.addNewHlink(), 0x00, 0x00, 0xff);
		setCTColor(clrScheme.addNewFolHlink(), 0x80, 0x00, 0x80);

		// Set default font scheme.
		CTFontScheme fontScheme = theme.getTheme().getThemeElements().addNewFontScheme();
		// Major fonts
		fontScheme.addNewMajorFont().addNewLatin().setTypeface("Cambria");
		fontScheme.getMajorFont().addNewEa().setTypeface("");
		fontScheme.getMajorFont().addNewCs().setTypeface("");

		// Minor fonts
		fontScheme.addNewMinorFont().addNewLatin().setTypeface("Calibri");
		fontScheme.getMinorFont().addNewEa().setTypeface("");
		fontScheme.getMinorFont().addNewCs().setTypeface("");

		// Set default format scheme.
		CTStyleMatrix fmtScheme = theme.getTheme().getThemeElements().addNewFmtScheme();
		initializeFillStyleList(fmtScheme.addNewFillStyleLst());

		initializeLineStyleList(fmtScheme.addNewLnStyleLst());

		initializeEffectStyleList(fmtScheme.addNewEffectStyleLst());

		initializeBGFillStyleList(fmtScheme.addNewBgFillStyleLst());
	}

	private void initializeFillStyleList(CTFillStyleList fillList) {
		fillList.addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);

		CTGradientFillProperties gradFill = fillList.addNewGradFill();
		gradFill.setRotWithShape(true);
		CTGradientStop gs = gradFill.addNewGsLst().addNewGs();
		gs.setPos(0);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewTint().setVal(50000);
		gs.getSchemeClr().addNewSatMod().setVal(300000);

		gs = gradFill.getGsLst().addNewGs();
		gs.setPos(35000);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewTint().setVal(37000);
		gs.getSchemeClr().addNewSatMod().setVal(300000);

		gs = gradFill.getGsLst().addNewGs();
		gs.setPos(100000);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewTint().setVal(15000);
		gs.getSchemeClr().addNewSatMod().setVal(350000);

		gradFill.addNewLin().setAng(16200000);
		gradFill.getLin().setScaled(true);

		gradFill = fillList.addNewGradFill();
		gradFill.setRotWithShape(true);
	}

	private void initializeLineStyleList(CTLineStyleList lineList) {
		CTLineProperties ln = lineList.addNewLn();
		ln.setW(9525);
		ln.setCap(STLineCap.FLAT);
		ln.setCmpd(STCompoundLine.SNG);
		ln.setAlgn(STPenAlignment.CTR);
		CTSchemeColor clr = ln.addNewSolidFill().addNewSchemeClr();
		clr.setVal(STSchemeColorVal.PH_CLR);
		clr.addNewShade().setVal(95000);
		clr.addNewSatMod().setVal(105000);
		ln.addNewPrstDash().setVal(STPresetLineDashVal.SOLID);

		ln = lineList.addNewLn();
		ln.setW(25400);
		ln.setCap(STLineCap.FLAT);
		ln.setCmpd(STCompoundLine.SNG);
		ln.setAlgn(STPenAlignment.CTR);
		clr = ln.addNewSolidFill().addNewSchemeClr();
		clr.setVal(STSchemeColorVal.PH_CLR);
		ln.addNewPrstDash().setVal(STPresetLineDashVal.SOLID);

		ln = lineList.addNewLn();
		ln.setW(38100);
		ln.setCap(STLineCap.FLAT);
		ln.setCmpd(STCompoundLine.SNG);
		ln.setAlgn(STPenAlignment.CTR);
		clr = ln.addNewSolidFill().addNewSchemeClr();
		clr.setVal(STSchemeColorVal.PH_CLR);
		ln.addNewPrstDash().setVal(STPresetLineDashVal.SOLID);
	}

	private void initializeEffectStyleList(CTEffectStyleList effectStyleList) {
		CTEffectStyleItem effectItem = effectStyleList.addNewEffectStyle();
		CTEffectList effectList = effectItem.addNewEffectLst();
		CTOuterShadowEffect shdw = effectList.addNewOuterShdw();
		shdw.setBlurRad(40000);
		shdw.setDist(20000);
		shdw.setDir(5400000);
		shdw.setRotWithShape(false);
		CTSRgbColor srgb = shdw.addNewSrgbClr();
		srgb.setVal(new byte[] {0, 0, 0});
		srgb.addNewAlpha().setVal(38000);

		effectItem = effectStyleList.addNewEffectStyle();
		effectList = effectItem.addNewEffectLst();
		shdw = effectList.addNewOuterShdw();
		shdw.setBlurRad(40000);
		shdw.setDist(23000);
		shdw.setDir(5400000);
		shdw.setRotWithShape(false);
		srgb = shdw.addNewSrgbClr();
		srgb.setVal(new byte[] {0, 0, 0});
		srgb.addNewAlpha().setVal(35000);

		effectItem = effectStyleList.addNewEffectStyle();
		effectList = effectItem.addNewEffectLst();
		shdw = effectList.addNewOuterShdw();
		shdw.setBlurRad(40000);
		shdw.setDist(23000);
		shdw.setDir(5400000);
		shdw.setRotWithShape(false);
		srgb = shdw.addNewSrgbClr();
		srgb.setVal(new byte[] {0, 0, 0});
		srgb.addNewAlpha().setVal(35000);

		CTScene3D scene3D = effectItem.addNewScene3D();
		scene3D.addNewCamera().setPrst(STPresetCameraType.ORTHOGRAPHIC_FRONT);
		scene3D.getCamera().addNewRot().setLat(0);
		scene3D.getCamera().getRot().setLon(0);
		scene3D.getCamera().getRot().setRev(0);
		scene3D.addNewLightRig().setRig(STLightRigType.THREE_PT);
		scene3D.getLightRig().setDir(STLightRigDirection.T);
		scene3D.getLightRig().addNewRot().setLat(0);
		scene3D.getLightRig().getRot().setLon(0);
		scene3D.getLightRig().getRot().setRev(1200000);
		CTBevel bevel = effectItem.addNewSp3D().addNewBevelT();
		bevel.setW(63500);
		bevel.setH(25400);
	}

	private void initializeBGFillStyleList(CTBackgroundFillStyleList bgFillStyleList) {
		bgFillStyleList.addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);

		CTGradientFillProperties gradFill = bgFillStyleList.addNewGradFill();
		gradFill.setRotWithShape(true);
		CTGradientStop gs = gradFill.addNewGsLst().addNewGs();
		gs.setPos(0);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewTint().setVal(40000);
		gs.getSchemeClr().addNewSatMod().setVal(350000);

		gs = gradFill.getGsLst().addNewGs();
		gs.setPos(40000);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewTint().setVal(45000);
		gs.getSchemeClr().addNewShade().setVal(99000);
		gs.getSchemeClr().addNewSatMod().setVal(350000);

		gs = gradFill.getGsLst().addNewGs();
		gs.setPos(100000);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewShade().setVal(20000);
		gs.getSchemeClr().addNewSatMod().setVal(255000);

		gradFill.addNewPath().setPath(STPathShadeType.CIRCLE);
		gradFill.getPath().addNewFillToRect().setL(50000);
		gradFill.getPath().getFillToRect().setT(-80000);
		gradFill.getPath().getFillToRect().setR(50000);
		gradFill.getPath().getFillToRect().setB(180000);

		gradFill = bgFillStyleList.addNewGradFill();
		gradFill.setRotWithShape(true);
		gs = gradFill.addNewGsLst().addNewGs();
		gs.setPos(0);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewTint().setVal(80000);
		gs.getSchemeClr().addNewSatMod().setVal(300000);

		gs = gradFill.getGsLst().addNewGs();
		gs.setPos(100000);
		gs.addNewSchemeClr().setVal(STSchemeColorVal.PH_CLR);
		gs.getSchemeClr().addNewShade().setVal(30000);
		gs.getSchemeClr().addNewSatMod().setVal(200000);

		gradFill.addNewPath().setPath(STPathShadeType.CIRCLE);
		gradFill.getPath().addNewFillToRect().setL(50000);
		gradFill.getPath().getFillToRect().setT(50000);
		gradFill.getPath().getFillToRect().setR(50000);
		gradFill.getPath().getFillToRect().setB(50000);
	}

	private void setCTColor(CTColor ctColor, int r, int g, int b) {
		ctColor.addNewSrgbClr().setVal(new byte[] {(byte) r, (byte) g, (byte) b});
	}

	@Override
	protected void commit() throws IOException {
		PackagePart part = getPackagePart();
		OutputStream out = part.getOutputStream();
		writeTo(out);
		out.close();
	}

	private void writeTo(OutputStream out) throws IOException {
		theme.save(out, new XmlOptions(DEFAULT_XML_OPTIONS));
	}
}
