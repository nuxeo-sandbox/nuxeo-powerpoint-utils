package nuxeo.powerpoint.utils.aspose;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;

import com.aspose.slides.IDocumentProperties;
import com.aspose.slides.IFontData;
import com.aspose.slides.ISlideCollection;
import com.aspose.slides.ISlideSize;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;

import nuxeo.powerpoint.utils.api.PowerPointUtils;

public class PowerPointUtilsWithAspose implements PowerPointUtils {

    public static final DateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS");
    private static final boolean IFontData = false;

    @Override
    public JSONObject getProperties(Blob blob) {
        
        JSONObject obj = new JSONObject();
        
        try {
            Presentation pres = new Presentation(blob.getStream());
            
            IDocumentProperties dp = pres.getDocumentProperties();
            
            ISlideSize slideSize = pres.getSlideSize();
            obj.put("Width", slideSize.getSize().getWidth());
            obj.put("Height", slideSize.getSize().getHeight());
            /*
            obj.put("AutoCompressPictures", ppt.getCTPresentation().getAutoCompressPictures());
            obj.put("CompatMode", ppt.getCTPresentation().getCompatMode());
            */

            // ================================== Properties
            obj.put("Category", dp.getCategory());
            obj.put("ContentStatus", dp.getContentStatus());
            obj.put("ContentType", dp.getContentType());
            obj.put("Created", DATE_FORMAT.format(dp.getCreatedTime()));
            obj.put("Creator", dp.getAuthor());
            //obj.put("Description", dp.getDescription());
            //obj.put("Identifier", dp.getIdentifier());
            obj.put("Keywords", dp.getKeywords());
            obj.put("LastModifiedByUser", dp.getLastSavedBy());
            obj.put("LastPrinted", DATE_FORMAT.format(dp.getLastPrinted()));
            obj.put("Modified", DATE_FORMAT.format(dp.getLastSavedTime()));
            obj.put("Revision", dp.getRevisionNumber());
            obj.put("Subject", dp.getSubject());
            obj.put("Title", dp.getTitle());
            // -------------------------------------
            obj.put("Application", dp.getNameOfApplication());
            obj.put("AppVersion", dp.getAppVersion());
            obj.put("Company", dp.getCompany());
            obj.put("HyperlinkBase", dp.getHyperlinkBase());
            obj.put("Manager", dp.getManager());
            obj.put("PresentationFormat", dp.getPresentationFormat());
            //obj.put("Template", dp.getTemplate());
            
            JSONArray arr = new JSONArray();
            for(IFontData font : pres.getFontsManager().getFonts()) {
                arr.put(font.getFontName());
            }
            obj.put("Fonts",  arr);

            
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

        File originalFile = blob.getFile();
        
        try {
            Presentation pres = new Presentation(blob.getStream());
            
            File tempDirectory = FileUtils.getTempDirectory();
            int slidesCount = pres.getSlides().size();
            for (int i = 0; i < slidesCount; i++) {
                
                Presentation destPres = new Presentation();
                ISlideCollection slds = destPres.getSlides();
                slds.addClone(pres.getSlides().get_Item(i));
                
                File newFile = new File(tempDirectory, fileNameBase + (i + 1) + ".pptx");
                destPres.save(newFile.getAbsolutePath(), SaveFormat.Pptx);
                FileBlob fb = new FileBlob(newFile.getAbsoluteFile());
                result.add(fb);
            }
            
        }catch (IOException e) {
            throw new NuxeoException(e);
        }
        
        return result;
    }

    @Override
    public BlobList splitPresentation(DocumentModel input, String xpath) throws IOException {
        // TODO Auto-generated method stub
        // return null;
        throw new UnsupportedOperationException();
    }

}
