/*
 * (C) Copyright 2020 Nuxeo (http://nuxeo.com/) and others.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * Contributors:
 *     Thibaud Arguillere
 */
package nuxeo.powerpoint.utils.test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

import jakarta.inject.Inject;

import org.json.JSONObject;
import org.junit.Ignore;
import org.junit.Test;
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

import nuxeo.powerpoint.utils.operations.GetPresentationPropertiesOp;
import nuxeo.powerpoint.utils.operations.GetSlideOp;
import nuxeo.powerpoint.utils.operations.GetThumbnailsOp;
import nuxeo.powerpoint.utils.operations.MergePresentationsOp;
import nuxeo.powerpoint.utils.operations.SplitPresentationOp;

/**
 * TODO: Not working since moving to LTS2023 a  nd Aspose 24.9,
 * and we need the plugin available quickly for other things
 * => To be explored "later"...
 * 
 * @since 10.10
 */
/*
 * We don't run a full test on the results => these are done in TestPowerPointUtilsWithApachePOI and
 * TestPowerPointUtilsWithAspose
 */
@RunWith(FeaturesRunner.class)
@Features(AutomationFeature.class)
@RepositoryConfig(init = DefaultRepositoryInit.class, cleanup = Granularity.METHOD)
@Deploy("nuxeo.powerpoint.utils-core")
public class TestOperations {

    @Inject
    protected CoreSession session;

    @Inject
    protected AutomationService automationService;

    @Test
    public void shouldSplitBlob() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);

        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID);

        assertNotNull(blobs);
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, blobs.size());

    }

    @Test
    @Ignore
    public void shouldSplitBlobWithAspose() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);
        Map<String, Object> params = new HashMap<>();
        params.put("useAspose", true);
        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID, params);

        assertNotNull(blobs);
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, blobs.size());

    }

    @Test
    public void shouldSplitDocument() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        DocumentModel doc = session.createDocumentModel("/", "pres", "File");
        doc.setPropertyValue("dc:title", "test-pres");
        doc.setPropertyValue("file:content", (Serializable) testFileBlob);
        doc = session.createDocument(doc);
        session.save();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(doc);
        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID);

        assertNotNull(blobs);
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, blobs.size());
    }

    @Test
    @Ignore
    public void shouldSplitDocumentWithAspose() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        DocumentModel doc = session.createDocumentModel("/", "pres", "File");
        doc.setPropertyValue("dc:title", "test-pres");
        doc.setPropertyValue("file:content", (Serializable) testFileBlob);
        doc = session.createDocument(doc);
        session.save();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(doc);
        Map<String, Object> params = new HashMap<>();
        params.put("useAspose", true);
        BlobList blobs = (BlobList) automationService.run(ctx, SplitPresentationOp.ID, params);

        assertNotNull(blobs);
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, blobs.size());
    }

    @Test
    // Merge uses Aspose
    @Ignore
    public void shouldMerge() throws Exception {

        BlobList blobs = new BlobList();

        File fileMerge1 = FileUtils.getResourceFileFromContext("files/merge1.pptx");
        Blob blob1 = new FileBlob(fileMerge1);
        blobs.add(blob1);
        File fileMerge2 = FileUtils.getResourceFileFromContext("files/merge2.pptx");
        Blob blob2 = new FileBlob(fileMerge2);
        blobs.add(blob2);
        File fileMerge3 = FileUtils.getResourceFileFromContext("files/merge3.pptx");
        Blob blob3 = new FileBlob(fileMerge3);
        blobs.add(blob3);

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(blobs);
        Blob result = (Blob) automationService.run(ctx, MergePresentationsOp.ID);

        assertNotNull(result);

    }

    @Test
    public void shouldGetProperties() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);

        String resultStr = (String) automationService.run(ctx, GetPresentationPropertiesOp.ID);
        assertNotNull(resultStr);
        // Just check no error is thrown
        @SuppressWarnings("unused")
        JSONObject resulJson = new JSONObject(resultStr);
        // Check that properties are valid are tested in TestPowerPointUtilsWithApachePOI and
        // TestPowerPointUtilsWithAspose
    }
    
    @Test
    public void shouldGetOneSlide() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);
        Map<String, Object> params = new HashMap<>();
        params.put("slideNumber", 4);

        Blob result = (Blob) automationService.run(ctx, GetSlideOp.ID, params);
        
        assertNotNull(result);
        // First slide is numbered 1, not zero (see PowerPointUtils interface)
        assertTrue(result.getFilename().endsWith("-5.pptx"));
    }
    
    @Test
    @Ignore
    public void shouldGetOneSlideWithAspose() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);
        Map<String, Object> params = new HashMap<>();
        params.put("useAspose", true);
        params.put("slideNumber", 4);

        Blob result = (Blob) automationService.run(ctx, GetSlideOp.ID, params);
        
        assertNotNull(result);
        // First slide is numbered 1, not zero (see PowerPointUtils interface)
        assertTrue(result.getFilename().endsWith("-5.pptx"));
    }
    
    @Test
    public void shouldGetThumbnails() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);

        BlobList blobs = (BlobList) automationService.run(ctx, GetThumbnailsOp.ID);
        
        assertNotNull(blobs);
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, blobs.size());
    }
    
    @Test
    @Ignore
    public void shouldGetThumbnailsWithAspose() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        OperationContext ctx = new OperationContext(session);
        ctx.setInput(testFileBlob);
        Map<String, Object> params = new HashMap<>();
        params.put("useAspose", true);

        BlobList blobs = (BlobList) automationService.run(ctx, GetThumbnailsOp.ID, params);
        
        assertNotNull(blobs);
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, blobs.size());
    }

}
