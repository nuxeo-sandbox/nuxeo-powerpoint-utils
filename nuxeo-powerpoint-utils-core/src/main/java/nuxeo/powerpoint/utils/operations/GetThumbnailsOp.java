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
@Operation(id = GetThumbnailsOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Get Thumbnails", description = "return a BlobList of thumbnails, one/slide."
        + " format can be \"jpg\" or \"png\"."
        + " maxWidth allows for returning smaller images. Any value <= 0 returns the images in the original dimension."
        + " If onlyVisible is true, thumbnails are returned only for visible slides.")
public class GetThumbnailsOp {

    public static final String ID = "Conversion.GetPowerPointThumbnails";

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath;

    @Param(name = "useAspose", required = false)
    protected Boolean useAspose = false;

    @Param(name = "maxWidth", required = false)
    protected Integer maxWidth = 0;

    @Param(name = "format", widget = Constants.W_OPTION, required = false, values = { "jpeg", "png" })
    protected String format = "png";

    @Param(name = "onlyVisible", required = false)
    protected Boolean onlyVisible = false;

    @OperationMethod
    public BlobList run(DocumentModel doc) throws IOException {

        BlobList result;

        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.getThumbnails(doc, xpath, maxWidth, format, onlyVisible);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.getThumbnails(doc, xpath, maxWidth, format, onlyVisible);
        }

        return result;
    }

    @OperationMethod
    public BlobList run(Blob blob) throws IOException {

        BlobList result;

        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.getThumbnails(blob, maxWidth, format, onlyVisible);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.getThumbnails(blob, maxWidth, format, onlyVisible);
        }

        return result;
    }
}
