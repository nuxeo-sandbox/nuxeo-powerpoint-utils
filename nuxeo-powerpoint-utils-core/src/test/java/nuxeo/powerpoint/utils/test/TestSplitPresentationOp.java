package nuxeo.powerpoint.utils.test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import javax.inject.Inject;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.nuxeo.common.utils.FileUtils;
import org.nuxeo.ecm.automation.AutomationService;
import org.nuxeo.ecm.automation.OperationContext;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.automation.test.AutomationFeature;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.ecm.core.test.DefaultRepositoryInit;
import org.nuxeo.ecm.core.test.annotations.Granularity;
import org.nuxeo.ecm.core.test.annotations.RepositoryConfig;
import org.nuxeo.runtime.test.runner.Deploy;
import org.nuxeo.runtime.test.runner.Features;
import org.nuxeo.runtime.test.runner.FeaturesRunner;

import nuxeo.powerpoint.utils.operations.SplitPresentationOp;

@RunWith(FeaturesRunner.class)
@Features(AutomationFeature.class)
@RepositoryConfig(init = DefaultRepositoryInit.class, cleanup = Granularity.METHOD)
@Deploy("nuxeo.powerpoint.utils-core")
public class TestSplitPresentationOp {

    public static final String BIG_PRESENTATION = "files/2020-Nuxeo-Overview-abstract.pptx";

    @Inject
    protected CoreSession session;

    @Inject
    protected AutomationService automationService;

    @Test
    public void shouldCallTheOperation() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
       
        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);
        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID);
        
        assertNotNull(blobs);
        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {

            assertEquals(fullPres.getSlides().size(), blobs.size());

            List<XSLFSlide> allSlides = fullPres.getSlides();

            for (int i = 0; i < blobs.size(); i++) {
                Blob blob = blobs.get(i);
                try (FileInputStream is = new FileInputStream(blob.getFile())) {
                    try (XMLSlideShow oneSlidePres = new XMLSlideShow(blob.getStream())) {
                        // Check we have only one
                        assertEquals(1, oneSlidePres.getSlides().size());

                        // Check the slides are the same
                        XSLFSlide originalSlide = allSlides.get(i);
                        XSLFSlide thisSlide = oneSlidePres.getSlides().get(0);
                        assertTrue(TestUtils.slidesLookTheSame(originalSlide, thisSlide));
                    }
                }
            }
        }
        
    }
}
