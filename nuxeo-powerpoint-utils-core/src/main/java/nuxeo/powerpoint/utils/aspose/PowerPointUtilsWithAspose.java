package nuxeo.powerpoint.utils.aspose;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
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

import com.aspose.slides.IAutoShape;
import com.aspose.slides.IFontData;
import com.aspose.slides.ILayoutSlide;
import com.aspose.slides.IMasterLayoutSlideCollection;
import com.aspose.slides.IMasterSlide;
import com.aspose.slides.IMasterSlideCollection;
import com.aspose.slides.IParagraph;
import com.aspose.slides.IShape;
import com.aspose.slides.ISlide;
import com.aspose.slides.ISlideCollection;
import com.aspose.slides.ITextFrame;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;

import freemarker.template.TemplateException;
import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;
import nuxeo.powerpoint.utils.api.PowerPointUtils;

/**
 * TODO: Not working since moving to LTS2023 a  nd Aspose 24.9,
 * and we need the plugin available quickly for other things
 * => To be explored "later"...
 * 
 * @since TODO
 */
public class PowerPointUtilsWithAspose implements PowerPointUtils {
    
    /*
     * TODO Fix the ClassDefNotFound for com.aspose.slides.Presentation.Presentation
     * and other errors.
     */
    public PowerPointUtilsWithAspose() {
        throw new NuxeoException("Version 2023.n of Nuxeo PowerPoint Utilities does not support Aspose.");
    }

    // ============================================================
    // PROPERTIES
    // ============================================================
    /*
     * ApachePOI actually can get more info than Aspose, especially in the statistics part
     * (number of words, f paragraphs, ...), as displayed to a user when using PowerPoint.
     * But ApacheOI (in the version we use, at least) does not have info on fonts.
     * So, basically, we get the info from Apache POI and add whatever is missing
     */
    @Override
    public JSONObject getProperties(Blob blob) {

        // Get everything Apache POI has in this implementation
        PowerPointUtilsWithApachePOI pptPoi = new PowerPointUtilsWithApachePOI();
        JSONObject obj = pptPoi.getProperties(blob);

        try {
            Presentation pres = new Presentation(blob.getStream());

            JSONArray arr;
            if (!obj.has("Fonts")) {
                arr = new JSONArray();
                obj.put("Fonts", arr);
            } else {
                arr = obj.getJSONArray("Fonts");
            }
            if (obj.getJSONArray("Fonts").length() == 0) {
                for (IFontData font : pres.getFontsManager().getFonts()) {
                    arr.put(font.getFontName());
                }
            }

        } catch (IOException | JSONException e) {
            throw new NuxeoException(e);
        }

        return obj;
    }

    // ============================================================
    // SPLIT
    // ============================================================
    @Override
    public BlobList splitPresentation(Blob blob) throws IOException {

        BlobList result = new BlobList();

        if (blob == null) {
            return result;
        }

        try {
            Presentation pres = new Presentation(blob.getStream());

            int slidesCount = pres.getSlides().size();
            for (int i = 0; i < slidesCount; i++) {

                Blob oneSlidePres = getSlide(blob, i);
                result.add(oneSlidePres);
            }

        } catch (IOException | NuxeoException e) {
            throw new NuxeoException(e);
        }

        return result;
    }

    @Override
    public BlobList splitPresentation(DocumentModel input, String xpath) throws IOException {

        BlobList blobs = splitPresentation(PowerPointUtils.getBlob(input, xpath));

        return blobs;
    }

    // ============================================================
    // MERGE
    // ============================================================
    @Override
    public Blob merge(BlobList blobs, boolean reuseMasters, String fileName) {

        Blob result = null;

        fileName = PowerPointUtils.checkMergedFileName(fileName);

        Presentation destPres = new Presentation();
        // May create a default slide, we want to start from scratch
        while (destPres.getSlides().size() > 0) {
            destPres.getSlides().removeAt(0);
        }
        destPres.getMasters().removeUnused(true);

        try {
            for (Blob b : blobs) {
                Presentation toMerge = new Presentation(b.getStream());
                if (toMerge != null) {
                    ISlideCollection slidesColl = toMerge.getSlides();
                    slidesColl.forEach(slide -> {

                        String slideTheme = slide.getLayoutSlide().getMasterSlide().getName();
                        String slideLayout = slide.getLayoutSlide().getName();

                        // TODO: Optimize _if needed_
                        // Benchmark and check if it would be really better to build/cache the master slides on the flow
                        IMasterSlide masterToUse = null;
                        if (reuseMasters) {
                            IMasterSlideCollection masterColl = destPres.getMasters();
                            for (int i = 0; i < masterColl.size(); i++) {
                                IMasterSlide master = masterColl.get_Item(i);
                                if (master != null && master.getName() != null && master.getName().equals(slideTheme)) {
                                    IMasterLayoutSlideCollection layoutMasterColl = master.getLayoutSlides();
                                    for (int j = 0; j < layoutMasterColl.size(); j++) {
                                        ILayoutSlide layoutMaster = layoutMasterColl.get_Item(j);
                                        if (layoutMaster.getName().equals(slideLayout)) {
                                            masterToUse = master;
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        if (masterToUse == null) {
                            destPres.getSlides().addClone(slide);
                        } else {
                            destPres.getSlides().addClone(slide, masterToUse, true);
                        }

                    });
                }
            }

            result = Blobs.createBlobWithExtension(".pptx");
            destPres.save(result.getFile().getAbsolutePath(), SaveFormat.Pptx);
            result.setFilename(fileName);
            result.setMimeType(PowerPointUtils.PPTX_MIMETYPE);

        } catch (IOException e) {
            throw new NuxeoException("Failed to merge PowerPoint persentations.", e);
        }

        return result;
    }

    @Override
    public Blob merge(DocumentModelList docs, String xpath, boolean reuseMasters, String fileName) {

        if (StringUtils.isBlank(xpath)) {
            xpath = "file:content";
        }

        BlobList blobs = new BlobList();
        for (DocumentModel doc : docs) {
            blobs.add(PowerPointUtils.getBlob(doc, xpath));
        }

        return merge(blobs, reuseMasters, fileName);
    }

    // ============================================================
    // GET SLIDE
    // ============================================================
    /*
     * Please see comment for splitPresentation() regarding Apache POI (need to duplicate the presentation
     * and delete all other slides)
     * Also check the interface : slideNumber is
     */
    @Override
    public Blob getSlide(Blob blob, int slideNumber) throws IOException {

        Blob result = null;

        if (blob == null) {
            return result;
        }

        String pptMimeType = PowerPointUtils.getBlobMimeType(blob);

        try {
            Presentation pres = new Presentation(blob.getStream());

            Presentation destPres = new Presentation();
            // May create a default slide, we want to start from scratch
            while (destPres.getSlides().size() > 0) {
                destPres.getSlides().removeAt(0);
            }
            destPres.getMasters().removeUnused(true);
            ISlideCollection slds = destPres.getSlides();

            slds.addClone(pres.getSlides().get_Item(slideNumber));

            result = Blobs.createBlobWithExtension(".pptx");
            destPres.save(result.getFile().getAbsolutePath(), SaveFormat.Pptx);

            // Update blob info
            result.setMimeType(pptMimeType);
            String fileNameBase = blob.getFilename();
            fileNameBase = FilenameUtils.getBaseName(fileNameBase);
            fileNameBase = StringUtils.appendIfMissing(fileNameBase, "-");
            // See interface: the file name must be 1-based, not zero-based
            result.setFilename(fileNameBase + (slideNumber + 1) + ".pptx");

        } catch (IOException e) {
            throw new NuxeoException("Failed to get slide #" + (slideNumber - 1), e);
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
            return result;
        }

        try {
            Presentation pres = new Presentation(blob.getStream());
            int slidesCount = pres.getSlides().size();
            for (int i = 0; i < slidesCount; i++) {
                ISlide slide = pres.getSlides().get_Item(i);
                if (onlyVisible && slide.getHidden()) {
                    continue;
                }
                
                Blob thumb = getThumbnail(slide, maxWidth, format);
                result.add(thumb);
            }

        } catch (IOException e) {
            throw new NuxeoException("Failed gerenate thumbnails.", e);
        }

        return result;

    }

    @Override
    public BlobList getThumbnails(DocumentModel doc, String xpath, int maxWidth, String format, boolean onlyVisible)
            throws IOException {

        BlobList blobs = getThumbnails(PowerPointUtils.getBlob(doc, xpath), maxWidth, format, onlyVisible);

        return blobs;
    }

    @Override
    public Blob getThumbnail(Blob blob, int slideNumber, int maxWidth, String format) throws IOException {

        Blob result = null;

        if (blob == null) {
            return result;
        }
        
        try {
            Presentation pres = new Presentation(blob.getStream());
            ISlide slide = pres.getSlides().get_Item(slideNumber);
            result = getThumbnail(slide, maxWidth, format);
            
        } catch (IOException e) {
            throw new NuxeoException("Failed gerenate thumbnails.", e);
        }
        
        return result;
        
    }

    @Override
    public Blob getThumbnail(DocumentModel doc, String xpath, int slideNumber, int maxWidth, String format) throws IOException {
        
        return getThumbnail(PowerPointUtils.getBlob(doc, xpath), slideNumber, maxWidth, format);
    }

    // ============================================================
    // OTHERS
    // ============================================================
    /**
     * Register Aspose with a valid license
     * See https://docs.aspose.com/display/slidesjava/Licensing
     * 
     * @param pathToLicenseFile
     * @since 10.10
     */
    public static void setLicense(String pathToLicenseFile) {
        com.aspose.slides.License license = new com.aspose.slides.License();
        license.setLicense(pathToLicenseFile);
    }

    // ============================================================
    // Replace Text
    // ============================================================
    // IMPORTANT RESTRICTION: The expression to replace (like ${doc["myschema:myfield"]})
    // must not contain lines.
    public Blob renderWithTemplate(DocumentModel doc, Blob template, String newFileName) throws Exception {

        Blob result = null;
        
        result = Blobs.createBlobWithExtension(".pptx");
        File resultFile = result.getFile();
        
        try {
            Presentation pres = new Presentation(template.getStream());
            ISlideCollection slides = pres.getSlides();

            for (ISlide slide : slides) {
                for (IShape shape : slide.getShapes()) {
                    if (shape instanceof IAutoShape) {
                        ITextFrame textFrame = ((IAutoShape) shape).getTextFrame();
                        if (textFrame != null) {
                            for (IParagraph para : textFrame.getParagraphs()) {
                                String text = para.getText();
                                String newText = replaceText(text, doc);
                                if(!newText.equals(text)) {
                                    para.setText(newText);
                                }
                                /*
                                for (IPortion portion : para.getPortions()) {
                                    if (portion.getText().contains(oldString)) {
                                        portion.setText(portion.getText().replace(oldString, newString));
                                    }
                                }
                                */
                            }
                        }
                    }
                }
            }

            pres.save(resultFile.getAbsolutePath(), SaveFormat.Pptx);

        } catch (IOException e) {
            throw new NuxeoException(e);
        }


        if(newFileName == null) {
            newFileName = template.getFilename();
        }
        if(newFileName == null) {
            newFileName = doc.getTitle();
        }
        if(!StringUtils.endsWithIgnoreCase(newFileName, ".pptx")) {
            newFileName += ".pptx";
        }
        result.setFilename(newFileName);
        result.setMimeType(PPTX_MIMETYPE);
        return result;
        
    }
    
    protected String replaceText(String text, DocumentModel doc) throws OperationException, RenderingException, TemplateException, IOException {
        OperationContext ctx = new OperationContext(doc.getCoreSession());
        ctx.setInput(doc);
        ctx.put("doc", doc);
        
        String newText = RenderingService.getInstance().render("ftl", text, ctx);

        return newText;
    }
    

    // ============================================================
    // Protected and specifics
    // ============================================================
    /*
     * Centralize getting a thnumbail once we have a slide
     */
    protected Blob getThumbnail(ISlide slide, int maxWidth, String format) throws IOException {
        
        Blob result = null;

        if (slide == null) {
            return result;
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
        
        double width = slide.getPresentation().getSlideSize().getSize().getWidth();
        double height = slide.getPresentation().getSlideSize().getSize().getHeight();
        
        float scale = 1;
        if (maxWidth > 0 && maxWidth < width) {
            scale = (float) (maxWidth / width);
            width = maxWidth;
            height = (int) (height * scale);
        }
        
        BufferedImage img = slide.getThumbnail(scale, scale);

        result = Blobs.createBlobWithExtension("." + format);
        javax.imageio.ImageIO.write(img, format, result.getFile());
        result.setMimeType(mimeType);
        // With Aspose, getSlideNumber() starts at 1, no need to (slide.getSlideNumber() + 1)
        result.setFilename("Slide " + slide.getSlideNumber() + "." + format);
        
        return result;
    }

}
