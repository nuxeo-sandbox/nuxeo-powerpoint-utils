package nuxeo.powerpoint.utils.operations;

import java.io.IOException;

import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;

import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;
import nuxeo.powerpoint.utils.aspose.PowerPointUtilsWithAspose;

/**
 *
 */
@Operation(id = GetSlideOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Get Slide", description = "Extract a slide from the input presentation."
        + " The blob will be named {original presentation name}-{slideNumberStartAt1}.pptx"
        + " slideNumber is the number of the slide. WARNING: It is zero-based, even if the output title starts at 1 (for better end user experience)."
        + " input can be a blob of the presentation, or a document. In this case xpath tells the operation which blob to use (file:content by default)."
        + " useAspose tells the operaiton to use Aspose for the rendition. Default is Apache POI. Slides rendered with Aspose have a better quality.")
public class GetSlideOp {

    public static final String ID = "Conversion.PowerPointGetSlide";

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath;

    @Param(name = "slideNumber", required = true)
    protected Integer slideNumber;

    @Param(name = "useAspose", required = false)
    protected Boolean useAspose = false;

    @OperationMethod
    public Blob run(DocumentModel doc) throws IOException {

        Blob result;
        
        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.getSlide(doc, xpath, slideNumber);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.getSlide(doc, xpath, slideNumber);
        }

        return result;
    }

    @OperationMethod
    public Blob run(Blob blob) throws IOException {
        Blob result;
        
        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.getSlide(blob, slideNumber);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.getSlide(blob, slideNumber);
        }

        return result;
    }
}
