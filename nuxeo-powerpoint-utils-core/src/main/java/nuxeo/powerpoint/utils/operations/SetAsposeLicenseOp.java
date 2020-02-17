package nuxeo.powerpoint.utils.operations;

import java.io.IOException;

import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;

import nuxeo.powerpoint.utils.aspose.PowerPointUtilsWithAspose;

/**
 *
 */
@Operation(id = SetAsposeLicenseOp.ID, category = Constants.CAT_CONVERSION, label = "PowerPoint: Set Aspose License", description = "Set Aspose license. Please visit https://docs.aspose.com/display/slidesjava/Licensing for details on the localtion of the license file.")
public class SetAsposeLicenseOp {

    public static final String ID = "Conversion.SetAsposeSlidesLicense";

    @Param(name = "licensePath", required = true)
    protected String licensePath;

    @OperationMethod
    public void run() throws IOException {

        PowerPointUtilsWithAspose.setLicense(licensePath);
    }
}
