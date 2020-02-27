package nuxeo.powerpoint.utils.operations;

import java.io.IOException;

import org.apache.commons.lang3.StringUtils;
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
@Operation(id = GetPresentationPropertiesOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Get Properties", description = "returns a JSON string containing information about the presentation, blob or document."
        + " If the input is a document, xpath is used (default to file:content).<br/>"
        + " We recommand to first try the operation and log the result to explore all the possible values.<br/>"
        + " When useAspose is true, the list of fonts is also returned.")
public class GetPresentationPropertiesOp {

    public static final String ID = "Conversion.PowerPointGetProperties";

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath = "file:content";

    @Param(name = "useAspose", required = false)
    protected Boolean useAspose = false;

    @OperationMethod
    public String run(DocumentModel doc) throws IOException {
        
        if (StringUtils.isBlank(xpath)) {
            xpath = "file:content";
        }
        Blob blob = (Blob) doc.getPropertyValue(xpath);

        return run(blob);
    }

    @OperationMethod
    public String run(Blob blob) throws IOException {

        if(useAspose) {
            PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
            return pptUtils.getProperties(blob).toString();
        } else {
            PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
            return pptUtils.getProperties(blob).toString();
        }
    }
}
