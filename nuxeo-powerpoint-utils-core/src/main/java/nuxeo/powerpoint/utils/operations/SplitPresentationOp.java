package nuxeo.powerpoint.utils.operations;

import java.io.IOException;

import org.apache.commons.lang3.StringUtils;
import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Context;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.PathRef;

import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;

/**
 *
 */
@Operation(id = SplitPresentationOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Split Presentation", description = "Split the input presentation"
        + " and returns an ordered list of blobs, one per slide."
        + " input can be a blob of the presentation, or a document. In this case xpath tells theoperaiton which blob to use (file:content by default)")
public class SplitPresentationOp {

    public static final String ID = "Conversion.SplitPPTOp";

    @Context
    protected CoreSession session;

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath;

    @OperationMethod
    public BlobList run(DocumentModel doc) throws IOException {

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();

        return pptUtils.splitPresentation(doc, xpath);
    }

    @OperationMethod
    public BlobList run(Blob blob) throws IOException {

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();

        return pptUtils.splitPresentation(blob);
    }
}
