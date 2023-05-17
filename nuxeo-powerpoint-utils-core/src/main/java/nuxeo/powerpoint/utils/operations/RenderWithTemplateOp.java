package nuxeo.powerpoint.utils.operations;

import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;

import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;
import nuxeo.powerpoint.utils.api.PowerPointUtils;
import nuxeo.powerpoint.utils.aspose.PowerPointUtilsWithAspose;

/**
 *
 */
@Operation(id = RenderWithTemplateOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Render Document with Template", description = ""
        + "Create a pptx from the template and the input doc."
        + " The template parameter is required, it is a blob holding a .pptx slides deck."
        + " Inside this template, add FreeMarker expressions, such as ${doc[\"schema:field\"]}"
        + " The operation replaces the values and returns a new blob."
        + " WARNING: an expression must be set on a single ligne. create a new text block in PowerPoint if needed."
        + " If fileName is empty, the returned blob will have the name of the template."
        + " useAspose tells the operation to use Aspose for the rendition. Default is Apache POI.")
public class RenderWithTemplateOp {

    public static final String ID = "Conversion.RenderDocumentWithPowerPointTemplate";

    @Param(name = "templateBlob", required = true)
    protected Blob templateBlob;

    @Param(name = "fileName", required = false)
    protected String fileName = null;

    @Param(name = "useAspose", required = false)
    protected Boolean useAspose = false;

    @OperationMethod
    public Blob run(DocumentModel doc) throws Exception {

        Blob result;
        
        PowerPointUtils pptUtils;

        if(useAspose) {
            pptUtils = new PowerPointUtilsWithAspose();
        } else {
            pptUtils = new PowerPointUtilsWithApachePOI();
        }
        result = pptUtils.renderWithTemplate(doc, templateBlob, fileName);

        return result;
    }
}
