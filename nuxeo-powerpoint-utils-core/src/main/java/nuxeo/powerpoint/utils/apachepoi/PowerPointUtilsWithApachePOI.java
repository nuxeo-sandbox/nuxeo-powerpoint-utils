/*
 * (C) Copyright 2020 Nuxeo (http://nuxeo.com/) and others.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * Contributors:
 *     Thibaud Arguillere
 */
package nuxeo.powerpoint.utils.apachepoi;

import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.ooxml.POIXMLProperties.ExtendedProperties;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.TextParagraph.FontAlign;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.usermodel.XSLFTheme;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.nuxeo.ecm.automation.OperationContext;
import org.nuxeo.ecm.automation.OperationException;
import org.nuxeo.ecm.automation.core.rendering.RenderingService;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.Blobs;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.DocumentModelList;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.ecm.platform.rendering.api.RenderingException;
import org.openxmlformats.schemas.presentationml.x2006.main.CTEmbeddedFontList;
import org.openxmlformats.schemas.presentationml.x2006.main.CTEmbeddedFontListEntry;

import freemarker.template.TemplateException;
import nuxeo.powerpoint.utils.api.PowerPointUtils;

/**
 * @since 10.10
 */
public class PowerPointUtilsWithApachePOI implements PowerPointUtils {

    public PowerPointUtilsWithApachePOI() {

    }

    // ============================================================
    // PROPERTIES
    // ============================================================
    @Override
    public JSONObject getProperties(Blob blob) {

        JSONObject obj = new JSONObject();

        try (XMLSlideShow ppt = new XMLSlideShow(blob.getStream())) {
            Dimension dim = ppt.getPageSize();
            obj.put("Width", dim.width);
            obj.put("Height", dim.height);

            obj.put("AutoCompressPictures", ppt.getCTPresentation().getAutoCompressPictures());
            obj.put("CompatMode", ppt.getCTPresentation().getCompatMode());

            // ================================== Properties
            POIXMLProperties props = ppt.getProperties();
            CoreProperties coreProps = props.getCoreProperties();
            obj.put("Category", coreProps.getCategory());
            obj.put("ContentStatus", coreProps.getContentStatus());
            obj.put("ContentType", coreProps.getContentType());
            obj.put("Created", DATE_FORMAT.format(coreProps.getCreated()));
            obj.put("Creator", coreProps.getCreator());
            obj.put("Description", coreProps.getDescription());
            obj.put("Identifier", coreProps.getIdentifier());
            obj.put("Keywords", coreProps.getKeywords());
            obj.put("LastModifiedByUser", coreProps.getLastModifiedByUser());
            obj.put("LastPrinted", DATE_FORMAT.format(coreProps.getLastPrinted()));
            obj.put("Modified", DATE_FORMAT.format(coreProps.getModified()));
            obj.put("Revision", coreProps.getRevision());
            obj.put("Subject", coreProps.getSubject());
            obj.put("Title", coreProps.getTitle());

            ExtendedProperties extProps = props.getExtendedProperties();
            obj.put("CountCharacters", extProps.getCharacters());
            obj.put("CountHiddenSlides", extProps.getHiddenSlides());
            obj.put("CountLines", extProps.getLines());
            obj.put("CountMMClips", extProps.getMMClips());
            obj.put("CountNotes", extProps.getNotes());
            obj.put("CountPages", extProps.getPages());
            obj.put("CountParagraphs", extProps.getParagraphs());
            obj.put("CountSlides", ppt.getSlides().size());
            obj.put("CountTotalTime", extProps.getTotalTime());
            obj.put("CountWords", extProps.getWords());
            // ----------------------------
            obj.put("Application", extProps.getApplication());
            obj.put("AppVersion", extProps.getAppVersion());
            obj.put("Company", extProps.getCompany());
            obj.put("HyperlinkBase", extProps.getHyperlinkBase());
            obj.put("Manager", extProps.getManager());
            obj.put("PresentationFormat", extProps.getPresentationFormat());
            obj.put("Template", extProps.getTemplate());

            // ================================== Objects and slides
            // Slides: Misc. info (title, master, theme, ...)
            JSONArray arr = new JSONArray();
            for (XSLFSlide slide : ppt.getSlides()) {
                JSONObject slideInfo = new JSONObject();
                slideInfo.put("SlideNumber", slide.getSlideNumber());
                slideInfo.put("Title", slide.getTitle() == null ? "" : slide.getTitle());
                slideInfo.put("Theme", slide.getTheme().getName());
                slideInfo.put("Master", slide.getSlideLayout().getName());
                arr.put(slideInfo);
            }
            obj.put("Slidesinfo", arr);

            // Master slides (array of themes, for each them, list of layouts)
            arr = new JSONArray();
            for (XSLFSlideMaster master : ppt.getSlideMasters()) {
                XSLFTheme masterThemeObj = master.getTheme();
                String masterTheme = masterThemeObj.getName();
                JSONArray arrLayouts = new JSONArray();
                for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                    arrLayouts.put(layout.getName());
                }
                JSONObject oneTheme = new JSONObject();
                oneTheme.put("Name", masterTheme);
                oneTheme.put("Layouts", arrLayouts);
                oneTheme.put("MasterFont", masterThemeObj.getMajorFont());
                oneTheme.put("MinorFont", masterThemeObj.getMinorFont());
                arr.put(oneTheme);
            }
            obj.put("MasterSlides", arr);

            // Embedded Fonts
            CTEmbeddedFontList fontList = ppt.getCTPresentation().getEmbeddedFontLst();
            arr = new JSONArray();
            if (fontList != null) {
                for (CTEmbeddedFontListEntry entry : fontList.getEmbeddedFontList()) {
                    arr.put(entry.getFont().getTypeface());
                }
            }
            obj.put("EmbeddedFonts", arr);

            // Nop easy way to get fonts with Apache POI 3.17
            // TODO: get fonts when switching to a higher version, if available
            arr = new JSONArray();
            obj.put("Fonts", arr);

        } catch (IOException | JSONException e) {
            throw new NuxeoException("Failed to get slides deck properties", e);
        }

        return obj;
    }

    // ============================================================
    // SPLIT
    // ============================================================
    /*
     * <b>IMPORTANT</b>
     * See {@link #getSlide()} for limitation and potential slowness. Basically, for each slide, we delete all the other
     * slides. This is surely a slow operation using CPU and i/o.
     */
    @Override
    public BlobList splitPresentation(Blob blob) throws IOException {

        BlobList result = new BlobList();

        if (blob == null) {
            return result;
        }

        try (XMLSlideShow ppt = new XMLSlideShow(blob.getStream())) {

            int slidesCount = ppt.getSlides().size();
            for (int i = 0; i < slidesCount; i++) {

                Blob oneSlidePres = getSlide(blob, i);
                result.add(oneSlidePres);
            }
        }

        return result;
    }

    /*
     * See <b>IMPORTANT</b> comment in <code>splitPresentation(Blob blob)</code>
     */
    @Override
    public BlobList splitPresentation(DocumentModel input, String xpath) throws IOException {

        BlobList blobs = splitPresentation(PowerPointUtils.getBlob(input, xpath));

        return blobs;
    }

    // ============================================================
    // MERGE
    // ============================================================
    // Unfortunately, merging with Apache POI is extremely complex as soon as a slide is complex
    // (complex background, specific font(s), multimedia file(s), ...)
    // Notice that merging can be done using Aspose instead.
    @Override
    public Blob merge(BlobList blobs, boolean reuseMasters, String fileName) {

        throw new UnsupportedOperationException("Merging slides is not supported with Apache POI, use Aspose instead.");
    }

    @Override
    public Blob merge(DocumentModelList docs, String xpath, boolean reuseMasters, String fileName) {

        throw new UnsupportedOperationException("Merging slides is not supported with Apache POI, use Aspose instead.");
    }

    // ============================================================
    // GET SLIDE
    // ============================================================
    /*
     * As of today (Apache POI 4.1.1, January 2020), it is very very difficult to extract a slide. There is no API for
     * this. Getting the master, the layout, getting the related/embedded images, graphs etc. is super cumbersome, and
     * for once, the Internet is not helping. It confirms it's extremely hard to make sure you extract a slide with all
     * its dependencies (master, layouts, images, videos, ...) so we are doing it in a different way: basically, for
     * each slide, we delete all the other slides. This is surely a slow operation using CPU and i/o.
     */
    @Override
    public Blob getSlide(Blob blob, int slideNumber) throws IOException {

        Blob result = null;

        if (blob == null) {
            return result;
        }

        String pptMimeType = PowerPointUtils.getBlobMimeType(blob);
        File originalFile = blob.getFile();

        try (XMLSlideShow ppt = new XMLSlideShow(blob.getStream())) {

            // Sanity check
            if (slideNumber < 0 || slideNumber >= ppt.getSlides().size()) {
                throw new NuxeoException("Invalid slide number: " + slideNumber);
            }

            // 1. Duplicate the original presentation
            result = Blobs.createBlobWithExtension(".pptx");
            File newFile = result.getFile();
            FileUtils.copyFile(originalFile, newFile);

            // 2. Remove slides in the copy
            try (InputStream is = new FileInputStream(newFile)) {
                try (XMLSlideShow copy = new XMLSlideShow(is)) {

                    // Remove slide(s) before
                    for (int iBefore = 0; iBefore < slideNumber; iBefore++) {
                        copy.removeSlide(0);
                    }

                    // Now, our slide is the first one
                    int tempSlidesCount = copy.getSlides().size();
                    for (int iAfter = 1; iAfter < tempSlidesCount; iAfter++) {
                        copy.removeSlide(1);
                    }

                    // Flush
                    try (FileOutputStream out = new FileOutputStream(newFile)) {
                        copy.write(out);
                    }

                    // Update blob info
                    result.setMimeType(pptMimeType);
                    String fileNameBase = blob.getFilename();
                    fileNameBase = FilenameUtils.getBaseName(fileNameBase);
                    fileNameBase = StringUtils.appendIfMissing(fileNameBase, "-");
                    // See interface: the file name must be 1-based, not zero-based
                    result.setFilename(fileNameBase + (slideNumber + 1) + ".pptx");
                }
            }

        }

        return result;

    }

    public Blob getSlide(DocumentModel input, String xpath, int slideNumber) throws IOException {

        return getSlide(PowerPointUtils.getBlob(input, xpath), slideNumber);
    }

    // ============================================================
    // THUMBNAILS
    // ============================================================
    @Override
    public BlobList getThumbnails(Blob blob, int maxWidth, String format, boolean onlyVisible) throws IOException {

        BlobList result = new BlobList();

        if (blob == null) {
            return null;
        }

        try (XMLSlideShow ppt = new XMLSlideShow(blob.getStream())) {
            for (XSLFSlide slide : ppt.getSlides()) {

                if (onlyVisible) {
                    // TODO This comes from Apache POI 4.n. directly use slide.isHidden() once Nuxeo platform has been
                    // upgraded
                    // (and this plugin also upgraded)
                    if (slide.getXmlObject().isSetShow() && !slide.getXmlObject().getShow()) {
                        continue;
                    }
                }

                Blob thumb = getThumbnail(slide, maxWidth, format);
                result.add(thumb);
            }
        }

        return result;
    }

    @Override
    public BlobList getThumbnails(DocumentModel doc, String xpath, int maxWidth, String format, boolean onlyVisible)
            throws IOException {

        if (StringUtils.isBlank(xpath)) {
            xpath = "file:content";
        }
        Blob blob = (Blob) doc.getPropertyValue(xpath);
        BlobList blobs = getThumbnails(blob, maxWidth, format, onlyVisible);

        return blobs;
    }

    @Override
    public Blob getThumbnail(Blob blob, int slideNumber, int maxWidth, String format) throws IOException {
        Blob result = null;

        if (blob == null) {
            return result;
        }

        try (XMLSlideShow ppt = new XMLSlideShow(blob.getStream())) {
            result = getThumbnail(ppt.getSlides().get(slideNumber), maxWidth, format);
        }

        return result;
    }

    @Override
    public Blob getThumbnail(DocumentModel doc, String xpath, int slideNumber, int maxWidth, String format)
            throws IOException {

        return getThumbnail(PowerPointUtils.getBlob(doc, xpath), slideNumber, maxWidth, format);
    }

    // ============================================================
    // OTHERS
    // ============================================================
    public Map<String, XSLFSlideMaster> getSlideMasters(XMLSlideShow slideShow) {

        HashMap<String, XSLFSlideMaster> namesAndMasters = new HashMap<String, XSLFSlideMaster>();
        for (XSLFSlideMaster master : slideShow.getSlideMasters()) {
            for (XSLFSlideLayout layout : master.getSlideLayouts()) {

                namesAndMasters.put(layout.getName(), master);
            }
        }

        return namesAndMasters;

    }

    // ============================================================
    // Replace Text
    // ============================================================
    // IMPORTANT RESTRICTIONS:
    // 1. The expression to replace (like ${doc["myschema:myfield"]})
    // must not contain lines.
    // 2. We handle only Freemarker expression starting with ${, and we expect an end }.
    public Blob renderWithTemplate(DocumentModel doc, Blob template, String newFileName) throws Exception {

        Blob result = null;

        File templateFile = template.getFile();

        result = Blobs.createBlobWithExtension(".pptx");
        File resultFile = result.getFile();

        try (InputStream is = new FileInputStream(templateFile.getAbsolutePath());
                OutputStream os = new FileOutputStream(resultFile.getAbsolutePath())) {

            try (XMLSlideShow ppt = new XMLSlideShow(is)) {
                for (XSLFSlide slide : ppt.getSlides()) {
                    for (XSLFShape shape : slide.getShapes()) {
                        if (shape instanceof XSLFTextShape) {
                            XSLFTextShape textShape = (XSLFTextShape) shape;

                            // We just check for freemarker indice, limiting the usage to ${ (so no <#if and loops etc.)
                            String textFromShape = textShape.getText();
                            if (textFromShape.indexOf("${") < 0) {
                                continue;
                            }

                            for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                                boolean isReplacing = false;
                                StringBuilder accumulatedText = new StringBuilder();
                                List<XSLFTextRun> involvedRuns = new ArrayList<>();

                                for (XSLFTextRun run : paragraph.getTextRuns()) {
                                    String textInRun = run.getRawText();
                                    if (!isReplacing && textInRun.contains("${")) {
                                        isReplacing = true;
                                    }
                                    if (isReplacing) {
                                        accumulatedText.append(textInRun);
                                        involvedRuns.add(run);
                                    }
                                    if (textInRun.contains("}")) {
                                        isReplacing = false;
                                        break;
                                    }
                                }

                                String finalAccumulated = accumulatedText.toString();
                                // System.out.println("finalAccumulated: " + finalAccumulated);
                                String replacedText = replaceText(finalAccumulated, doc);
                                // System.out.println("replacedText: " + replacedText);

                                if (!replacedText.equals(finalAccumulated)) {
                                    // Now distribute the replaced text back to the involved runs
                                    int textStart = 0;
                                    int remainingLength = replacedText.length(); // Length of the remaining part of the
                                                                                 // replaced text
                                    int runIndex;
                                    for (runIndex = 0; runIndex < involvedRuns.size(); runIndex++) {
                                        XSLFTextRun run = involvedRuns.get(runIndex);
                                        if (remainingLength > 0) {
                                            // If there's still some replacement text left
                                            int originalRunLength = run.getRawText().length();
                                            if (originalRunLength < remainingLength) {
                                                // If the original run's text is shorter than the remaining part of the
                                                // replaced text,
                                                // take as much of the replaced text as the length of the original run's
                                                // text
                                                String newTextForRun = replacedText.substring(textStart,
                                                        textStart + originalRunLength);
                                                run.setText(newTextForRun);
                                                textStart += originalRunLength;
                                                remainingLength -= originalRunLength;
                                            } else {
                                                // If the original run's text is not shorter than the remaining part of
                                                // the replaced text,
                                                // put the whole remaining part of the replaced text into this run
                                                String newTextForRun = replacedText.substring(textStart);
                                                run.setText(newTextForRun);
                                                remainingLength = 0; // No replacement text left
                                            }
                                        } else {
                                            // If there's no replacement text left but there are still some runs that
                                            // were part of the original placeholder,
                                            // set their text to an empty string
                                            run.setText("");
                                        }
                                    }

                                    // If there's still some replacement text left after going through all involved
                                    // runs, create new runs
                                    while (remainingLength > 0) {
                                        XSLFTextRun newRun = paragraph.addNewTextRun();
                                        String newTextForRun = replacedText.substring(textStart);
                                        newRun.setText(newTextForRun);
                                        remainingLength = 0;
                                    }
                                }
                            }

                            /*
                             * Oroginal code (first implementation)
                             * for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                             * for (XSLFTextRun run : paragraph.getTextRuns()) {
                             * String text = run.getRawText();
                             * String newText = replaceText(text, doc);
                             * if(!text.equals(newText)) {
                             * run.setText(newText);
                             * }
                             * }
                             * }
                             */
                        }
                    }
                }

                ppt.write(os);

            } catch (IOException e) {
                throw new NuxeoException(e);
            }
        } catch (IOException e) {
            throw new NuxeoException(e);
        }

        if (newFileName == null) {
            newFileName = template.getFilename();
        }
        if (newFileName == null) {
            newFileName = doc.getTitle();
        }
        if (!StringUtils.endsWithIgnoreCase(newFileName, ".pptx")) {
            newFileName += ".pptx";
        }
        result.setFilename(newFileName);
        result.setMimeType(PPTX_MIMETYPE);
        return result;
    }

    protected String replaceText(String text, DocumentModel doc)
            throws OperationException, RenderingException, TemplateException, IOException {
        OperationContext ctx = new OperationContext(doc.getCoreSession());
        ctx.setInput(doc);
        ctx.put("doc", doc);

        String newText = RenderingService.getInstance().render("ftl", text, ctx);

        return newText;
    }

    // ============================================================
    // PROTECTED AND SPECIFICS
    // ============================================================
    protected Blob getThumbnail(XSLFSlide slide, int maxWidth, String format) throws IOException {

        Blob result = null;

        if (slide == null) {
            return null;
        }

        if (StringUtils.isBlank(format)) {
            format = "png";
        }

        String mimeType;
        switch (format.toLowerCase()) {
        case "jpg":
        case "jpeg":
            format = "jpg";
            mimeType = "image/jpeg";
            break;

        case "png":
            mimeType = "image/png";
            break;

        default:
            throw new NuxeoException(format + " is no a supported formats (only jpg or png)");
        }

        Dimension pgsize = slide.getSlideShow().getPageSize();
        int width = pgsize.width;
        int height = pgsize.height;

        float scale = 1;
        if (maxWidth > 0 && maxWidth < width) {
            scale = (float) maxWidth / (float) width;
            width = maxWidth;
            height = (int) (height * scale);
        }

        // Thanks to Apache example, PPTX2PNG
        // BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = img.createGraphics();

        /*
         * graphics.setPaint(Color.white);
         * graphics.fill(new Rectangle2D.Float(0, 0, width, height));
         */

        // default rendering options
        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);

        graphics.scale(scale, scale);
        slide.draw(graphics);

        result = Blobs.createBlobWithExtension("." + format);
        javax.imageio.ImageIO.write(img, format, result.getFile());
        result.setMimeType(mimeType);
        // getSlideNumber() returns a number starting at 1 (as expected by a user)
        result.setFilename("Slide " + slide.getSlideNumber() + "." + format);

        return result;
    }

}
