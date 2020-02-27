package nuxeo.powerpoint.utils.operations;

import java.io.IOException;

import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;

import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;
import nuxeo.powerpoint.utils.aspose.PowerPointUtilsWithAspose;

/**
 *
 */
@Operation(id = GetThumbnailOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Get a Thumbnail", description = "return a Blob of thumbnail of the slide."
        + " slideNumber is the number of the slide, zero-based. WARNING: , even if the output title starts at 1 (for better end user experience)."
        + " format can be \"jpg\" or \"png\"."
        + " maxWidth allows for returning smaller images. Any value <= 0 returns the images in the original dimension."
        + " If onlyVisible is true, thumbnails are returned only for visible slides."
        + " useAspose tells the operaiton to use Aspose for the rendition. Default is Apache POI. Slides rendered with Aspose have a better quality.")
public class GetThumbnailOp {

    public static final String ID = "Conversion.PowerPointGetOneThumbnail";

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath;

    @Param(name = "slideNumber", required = true)
    protected Integer slideNumber;

    @Param(name = "maxWidth", required = false)
    protected Integer maxWidth = 0;

    @Param(name = "format", widget = Constants.W_OPTION, required = false, values = { "jpeg", "png" })
    protected String format = "png";

    @Param(name = "useAspose", required = false)
    protected Boolean useAspose = false;

    @OperationMethod
    public Blob run(DocumentModel doc) throws IOException {

        Blob result;

        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.getThumbnail(doc, xpath, slideNumber, maxWidth, format);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.getThumbnail(doc, xpath, slideNumber, maxWidth, format);
        }

        return result;
    }

    @OperationMethod
    public Blob run(Blob blob) throws IOException {

        Blob result;

        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.getThumbnail(blob, slideNumber, maxWidth, format);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.getThumbnail(blob, slideNumber, maxWidth, format);
        }

        return result;
    }
}
