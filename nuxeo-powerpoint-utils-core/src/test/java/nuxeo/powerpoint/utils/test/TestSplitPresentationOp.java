package nuxeo.powerpoint.utils.test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileInputStream;
import java.io.Serializable;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.inject.Inject;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;
import org.junit.runner.RunWith;
import org.nuxeo.common.utils.FileUtils;
import org.nuxeo.ecm.automation.AutomationService;
import org.nuxeo.ecm.automation.OperationContext;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.automation.test.AutomationFeature;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.DocumentModel;
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
    
    @Rule
    public TestName testName = new TestName();
    
    // Centralizing the check of the result of the split operation
    protected void checkSplitResult(Blob presentation, BlobList blobs) throws Exception {
        
        assertNotNull(testName.getMethodName() + " failed", blobs);
        
        try (XMLSlideShow fullPres = new XMLSlideShow(presentation.getStream())) {

            assertEquals(fullPres.getSlides().size(), blobs.size());

            List<XSLFSlide> allSlides = fullPres.getSlides();

            for (int i = 0; i < blobs.size(); i++) {
                Blob blob = blobs.get(i);
                try (FileInputStream is = new FileInputStream(blob.getFile())) {
                    try (XMLSlideShow oneSlidePres = new XMLSlideShow(blob.getStream())) {
                        // Check we have only one
                        assertEquals(testName.getMethodName() + " failed", 1, oneSlidePres.getSlides().size());

                        // Check the slides are the same
                        XSLFSlide originalSlide = allSlides.get(i);
                        XSLFSlide thisSlide = oneSlidePres.getSlides().get(0);
                        assertTrue(testName.getMethodName() + " failed",TestUtils.slidesLookTheSame(originalSlide, thisSlide));
                    }
                }
            }
        }
        
    }

    @Test
    public void shouldSplitTheBlob() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
       
        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);
        
        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID);
                
        checkSplitResult(testFileBlob, blobs);
    }

    @Test
    public void shouldSplitTheBlobWithAspose() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
       
        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);
        Map<String, Object> params = new HashMap<>();
        params.put("useAspose", true);
        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID, params);
        
        checkSplitResult(testFileBlob, blobs);
    }

    @Test
    public void shouldSplitTheDocument() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
       
        DocumentModel doc = session.createDocumentModel("/", "pres", "File");
        doc.setPropertyValue("dc:title", "test-pres");
        doc.setPropertyValue("file:content", (Serializable) testFileBlob);
        doc = session.createDocument(doc);
        session.save();
        
        OperationContext ctx = new OperationContext(session);
        ctx.setInput(doc);
        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID);
        
        checkSplitResult(testFileBlob, blobs);
    }
}
