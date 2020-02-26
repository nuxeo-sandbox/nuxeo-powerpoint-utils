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
@Operation(id = SplitPresentationOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Split Presentation", description = "Split the input presentation"
        + " and returns an ordered list of blobs, one per slide."
        + " input can be a blob of the presentation, or a document. In this case xpath tells the operation which blob to use (file:content by default)")
public class SplitPresentationOp {

    public static final String ID = "Conversion.SplitPowerPointPresentation";

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath;

    @Param(name = "useAspose", required = false)
    protected Boolean useAspose = false;

    @OperationMethod
    public BlobList run(DocumentModel doc) throws IOException {

        BlobList result;
        
        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.splitPresentation(doc, xpath);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.splitPresentation(doc, xpath);
        }

        return result;
    }

    @OperationMethod
    public BlobList run(Blob blob) throws IOException {
        BlobList result;
        if (useAspose) {
            PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
            result = asposePptUtils.splitPresentation(blob);
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            result = pptUtils.splitPresentation(blob);
        }

        return result;
    }
}
