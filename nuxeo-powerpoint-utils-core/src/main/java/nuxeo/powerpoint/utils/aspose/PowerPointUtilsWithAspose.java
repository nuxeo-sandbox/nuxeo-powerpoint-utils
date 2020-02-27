package nuxeo.powerpoint.utils.aspose;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.Blobs;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.DocumentModelList;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;

import com.aspose.slides.IFontData;
import com.aspose.slides.ILayoutSlide;
import com.aspose.slides.IMasterLayoutSlideCollection;
import com.aspose.slides.IMasterSlide;
import com.aspose.slides.IMasterSlideCollection;
import com.aspose.slides.ISlide;
import com.aspose.slides.ISlideCollection;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;

import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;
import nuxeo.powerpoint.utils.api.PowerPointUtils;

public class PowerPointUtilsWithAspose implements PowerPointUtils {

    // ==============================> PROPERTIES
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

        try {
            Presentation pres = new Presentation(blob.getStream());

            File tempDirectory = FileUtils.getTempDirectory();
            int slidesCount = pres.getSlides().size();
            for (int i = 0; i < slidesCount; i++) {

                Presentation destPres = new Presentation();
                // May create a default slide, we want to start from scratch
                while (destPres.getSlides().size() > 0) {
                    destPres.getSlides().removeAt(0);
                }
                destPres.getMasters().removeUnused(true);
                ISlideCollection slds = destPres.getSlides();
                slds.addClone(pres.getSlides().get_Item(i));

                File newFile = new File(tempDirectory, fileNameBase + (i + 1) + ".pptx");
                destPres.save(newFile.getAbsolutePath(), SaveFormat.Pptx);
                FileBlob fb = new FileBlob(newFile.getAbsoluteFile());
                fb.setMimeType(pptMimeType);
                result.add(fb);
            }

        } catch (IOException e) {
            throw new NuxeoException("Failed to split PowerPoint presentation.", e);
        }

        return result;
    }

    // ==============================> SPLIT
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
        for(DocumentModel doc : docs) {
            blobs.add((Blob) doc.getPropertyValue(xpath));
        }
        
        return merge(blobs, reuseMasters, fileName);
    }
    
    // ==============================> THUMBNAILS
    @Override
    public BlobList getThumbnails(Blob blob, int maxWidth, String format, boolean onlyVisible) throws IOException {

        BlobList result = new BlobList();

        if (blob == null) {
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
        
        try {
            Presentation pres = new Presentation(blob.getStream());
            
            double width = pres.getSlideSize().getSize().getWidth();
            double height = pres.getSlideSize().getSize().getHeight();

            float scale = 1;
            if (maxWidth > 0 && maxWidth < width) {
                scale = (float) (maxWidth / width);
                width = maxWidth;
                height = (int) (height * scale);
            }
            
            int slidesCount = pres.getSlides().size();
            for (int i = 0; i < slidesCount; i++) {

                ISlide slide = pres.getSlides().get_Item(i);
                
                if(onlyVisible && slide.getHidden()) {
                    continue;
                }
                
                BufferedImage img = slide.getThumbnail(scale, scale);
                
                Blob b = Blobs.createBlobWithExtension("." + format);
                javax.imageio.ImageIO.write(img, format, b.getFile());
                b.setMimeType(mimeType);

                b.setFilename("Slide " + (i + 1) + "." + format);
                result.add(b);
            }

        } catch (IOException e) {
            throw new NuxeoException("Failed gerenate thumbnails.", e);
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
    
    /**
     * Register Aspose with a valid license
     * 
     * See https://docs.aspose.com/display/slidesjava/Licensing
     * 
     * @param pathToLicenseFile
     * @since 10.10
     */
    public static void setLicense(String pathToLicenseFile) {
        com.aspose.slides.License license = new com.aspose.slides.License();
        license.setLicense(pathToLicenseFile);
    }

}
