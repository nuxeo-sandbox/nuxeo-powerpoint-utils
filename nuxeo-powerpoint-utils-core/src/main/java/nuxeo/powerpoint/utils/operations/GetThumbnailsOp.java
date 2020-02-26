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
        + " ignoreWhen Hidden is true, thumbnails are returned only for visible slides.")
public class GetThumbnailsOp {

    public static final String ID = "Conversion.GetPowerPointThumbnails";

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath;

    @Param(name = "useAspose", required = false)
    protected Boolean useAspose = false;

    @OperationMethod
    public BlobList run(DocumentModel doc) throws IOException {

        throw new UnsupportedOperationException("Not yet implemented");
    }

    @OperationMethod
    public BlobList run(Blob blob) throws IOException {

        throw new UnsupportedOperationException("Not yet implemented");
    }
}
