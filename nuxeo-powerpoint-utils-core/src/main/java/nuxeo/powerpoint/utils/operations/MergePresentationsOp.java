package nuxeo.powerpoint.utils.operations;

import java.io.IOException;

import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModelList;

import nuxeo.powerpoint.utils.aspose.PowerPointUtilsWithAspose;

/**
 *
 */
@Operation(id = MergePresentationsOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Merge Presentations", description = "Merge the input presentations"
        + " and returns the result.<br/>"
        + " input can be a list of blobs or documents. In this case xpath tells the operation which blob to use (file:content by default)."
        + " If fileName is empty, the result is named merged.pptx.<br/>"
        + " If reuseMasters is false, the whole set of master slides of each presentation to merge is copied to the destination."
        + " Else, they are copied only if the same masters (same theme, same layout) don't exist yet in the merged result.<br/>"
        + " IMPORTANT: This operation uses Aspose (aspose.com), which requires a valid license. Without a license all slides are watermarked.")
public class MergePresentationsOp {

    public static final String ID = "Conversion.PowerPointMerge";

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath;

    @Param(name = "fileName", required = false)
    protected String fileName = null;

    @Param(name = "reuseMasters", required = false)
    protected Boolean reuseMasters = false;

    @OperationMethod
    public Blob run(DocumentModelList docs) throws IOException {

        Blob result;

        PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
        result = asposePptUtils.merge(docs, xpath, reuseMasters, fileName);

        return result;
    }

    @OperationMethod
    public Blob run(BlobList blobs) throws IOException {
        
        Blob result;
        
        PowerPointUtilsWithAspose asposePptUtils = new PowerPointUtilsWithAspose();
        result = asposePptUtils.merge(blobs, reuseMasters, fileName);

        return result;
    }
}
