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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.POIXMLProperties.ExtendedProperties;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTheme;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.DocumentModelList;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.ecm.platform.mimetype.MimetypeDetectionException;
import org.nuxeo.ecm.platform.mimetype.MimetypeNotFoundException;
import org.nuxeo.ecm.platform.mimetype.interfaces.MimetypeRegistry;
import org.nuxeo.ecm.platform.mimetype.service.MimetypeRegistryService;
import org.nuxeo.runtime.api.Framework;
import org.openxmlformats.schemas.presentationml.x2006.main.CTEmbeddedFontList;
import org.openxmlformats.schemas.presentationml.x2006.main.CTEmbeddedFontListEntry;

import nuxeo.powerpoint.utils.api.PowerPointUtils;

/**
 * @since 10.10
 */
public class PowerPointUtilsWithApachePOI implements PowerPointUtils {

    public PowerPointUtilsWithApachePOI() {

    }

    // ==============================> PROPERTIES
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
            obj.put("CountSlides", extProps.getSlides());
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
                XSLFTheme  masterThemeObj = master.getTheme();
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

    // ==============================> SPLIT
    /*
     * <b>IMPORTANT</b>
     * As of today (Apache POI 4.1.1, January 2020), it is very very difficult to extract a slide. There is no API for
     * this. Getting the master, the layout, getting the related/embedded images, graphs etc. is super cumbersome, and
     * for once, the Internet is not helping. It confirms it's extremely hard to make sure you extract a slide with all
     * its dependencies (master, layouts, images, videos, ...) so we are doing it in a different way: basically, for
     * each slide, we delete all the other slides. This is surely a slow operation using CPU and i/o.
     */
    @Override
    public BlobList splitPresentation(Blob blob) throws IOException {

        BlobList result = new BlobList();
        String pptMimeType;
        String fileNameBase;

        if (blob == null) {
            return result;
        }

        pptMimeType = PowerPointUtils.getBlobMimeType(blob);
        fileNameBase = blob.getFilename();
        fileNameBase = FilenameUtils.getBaseName(fileNameBase);
        fileNameBase = StringUtils.appendIfMissing(fileNameBase, "-");

        File originalFile = blob.getFile();

        try (XMLSlideShow ppt = new XMLSlideShow(blob.getStream())) {

            File tempDirectory = FileUtils.getTempDirectory();
            int slidesCount = ppt.getSlides().size();
            for (int i = 0; i < slidesCount; i++) {

                // 1. Duplicate the original presentation
                File newFile = new File(tempDirectory, fileNameBase + (i + 1) + ".pptx");
                FileUtils.copyFile(originalFile, newFile);

                // 2. Remove slides in the copy
                try (InputStream is = new FileInputStream(newFile)) {
                    try (XMLSlideShow copy = new XMLSlideShow(is)) {

                        for (int iBefore = 0; iBefore < i; iBefore++) {
                            copy.removeSlide(0);
                        }

                        // Now, our slide is the first one (0 based)
                        int tempSlidesCount = copy.getSlides().size();
                        for (int iAfter = 1; iAfter < tempSlidesCount; iAfter++) {
                            copy.removeSlide(1);
                        }

                        try (FileOutputStream out = new FileOutputStream(newFile)) {
                            copy.write(out);
                        }
                    }
                }

                // 3. Save as blob
                FileBlob oneSlidePres = new FileBlob(newFile, pptMimeType);
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

        if (StringUtils.isBlank(xpath)) {
            xpath = "file:content";
        }
        Blob blob = (Blob) input.getPropertyValue(xpath);
        BlobList blobs = splitPresentation(blob);

        return blobs;
    }
    
    // ==============================> MERGE
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

    // ==============================> Utilities
    public Map<String, XSLFSlideMaster> getSlideMasters(XMLSlideShow slideShow) {

        HashMap<String, XSLFSlideMaster> namesAndMasters = new HashMap<String, XSLFSlideMaster>();
        for (XSLFSlideMaster master : slideShow.getSlideMasters()) {
            for (XSLFSlideLayout layout : master.getSlideLayouts()) {

                namesAndMasters.put(layout.getName(), master);
            }
        }

        return namesAndMasters;

    }

}
