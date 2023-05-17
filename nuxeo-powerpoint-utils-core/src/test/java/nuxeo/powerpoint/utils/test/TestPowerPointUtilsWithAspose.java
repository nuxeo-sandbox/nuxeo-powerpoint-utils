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
import static org.junit.Assert.assertNotEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import javax.inject.Inject;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Ignore;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.nuxeo.common.utils.FileUtils;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.automation.test.AutomationFeature;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.blobholder.BlobHolder;
import org.nuxeo.ecm.core.api.blobholder.SimpleBlobHolder;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.ecm.core.convert.api.ConversionService;
import org.nuxeo.ecm.core.test.DefaultRepositoryInit;
import org.nuxeo.ecm.core.test.annotations.Granularity;
import org.nuxeo.ecm.core.test.annotations.RepositoryConfig;
import org.nuxeo.ecm.platform.picture.api.ImageInfo;
import org.nuxeo.ecm.platform.picture.api.ImagingService;
import org.nuxeo.runtime.test.runner.Deploy;
import org.nuxeo.runtime.test.runner.Features;
import org.nuxeo.runtime.test.runner.FeaturesRunner;

import com.aspose.slides.Presentation;

import nuxeo.powerpoint.utils.aspose.PowerPointUtilsWithAspose;

/**
 * @since 10.10
 */
/*
 * For testing the merge we have 3 presentations (files/merge1.pptx etc.
 * The 2 first ones have different slides but the same masters.
 * The third one has different master slides
 * This impacts the test, depending on parameters passed to merge().
 */
@RunWith(FeaturesRunner.class)
@Features(AutomationFeature.class)
@RepositoryConfig(init = DefaultRepositoryInit.class, cleanup = Granularity.METHOD)
@Deploy({ "org.nuxeo.ecm.platform.picture.core", "org.nuxeo.ecm.platform.tag", "nuxeo.powerpoint.utils-core" })
public class TestPowerPointUtilsWithAspose {

    @Inject
    protected CoreSession session;

    @Inject
    protected ImagingService imagingService;
    
    @Inject
    protected ConversionService conversionService;

    @Test
    public void shouldSplitABlobPresentation() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        BlobList blobs = pptUtils.splitPresentation(testFileBlob);

        assertNotNull(blobs);
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, blobs.size());

        // For quick tests on your Mac :-)
        // for (Blob b : blobs) {
        // TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        // }

        // First slide is numbered 1, not zero (see PowerPointUtils interface)
        assertTrue(blobs.get(0).getFilename().endsWith("-1.pptx"));

        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {

            assertEquals(fullPres.getSlides().size(), blobs.size());

            List<XSLFSlide> allSlides = fullPres.getSlides();

            // We check with Apache POI.
            for (int i = 0; i < blobs.size(); i++) {
                Blob blob = blobs.get(i);
                try (FileInputStream is = new FileInputStream(blob.getFile())) {
                    try (XMLSlideShow oneSlidePres = new XMLSlideShow(blob.getStream())) {
                        // Check we have only one
                        // WARNING: If using Aspose in demo mode, we always have a "Built with Aspose slide"
                        int countSlides = oneSlidePres.getSlides().size();
                        assertTrue(countSlides == 1);

                        // Check the slides are the same
                        XSLFSlide originalSlide = allSlides.get(i);
                        XSLFSlide thisSlide = oneSlidePres.getSlides().get(countSlides - 1);
                        assertTrue(TestUtils.slidesLookTheSame(originalSlide, thisSlide));
                    }
                }
            }
        }

    }

    @Test
    public void tesGetProperties() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        JSONObject result = pptUtils.getProperties(testFileBlob);

        // System.out.println("\n" + result.toString(2));

        // See, in PowerPoint, File > Properties of the test file.
        assertEquals("Nuxeo Unit Testing", result.get("Creator"));
        assertEquals("Nuxeo", result.get("Company"));
        assertEquals("Widescreen", result.get("PresentationFormat"));

        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_SLIDES_COUNT, result.get("CountSlides"));
        assertEquals(TestUtils.MAIN_TEST_PRESENTATION_HIDDEN_SLIDES, result.get("CountHiddenSlides"));

        JSONArray arr = result.getJSONArray("MasterSlides");
        assertEquals(2, arr.length());
        // First one is "Office Theme"
        JSONObject theme = arr.getJSONObject(0);
        // getJSONObject does not return null is there is no value, it throws an exception
        assertEquals("Office Theme", theme.get("Name"));
        // Could also check the layouts...

        // Could also check info on every slides...
    }

    @Test
    public void shouldMergePresentationsWithCopyMasterSlides() throws Exception {

        BlobList blobs = new BlobList();

        // (See Class comments)
        File fileMerge1 = FileUtils.getResourceFileFromContext("files/merge1.pptx");
        Blob blob1 = new FileBlob(fileMerge1);
        blobs.add(blob1);
        File fileMerge2 = FileUtils.getResourceFileFromContext("files/merge2.pptx");
        Blob blob2 = new FileBlob(fileMerge2);
        blobs.add(blob2);
        File fileMerge3 = FileUtils.getResourceFileFromContext("files/merge3.pptx");
        Blob blob3 = new FileBlob(fileMerge3);
        blobs.add(blob3);

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        Blob resultBlob = pptUtils.merge(blobs, false, null);

        assertNotNull(resultBlob);

        // TestUtils.saveBlobOnDesktop(resultBlob, "test-ppt-utils");

        // We passed null as fileName => the code should provide the default name
        assertEquals("merged.pptx", resultBlob.getFilename());

        // Check slides
        Presentation src1 = new Presentation(blob1.getStream());
        Presentation src2 = new Presentation(blob2.getStream());
        Presentation src3 = new Presentation(blob3.getStream());
        int countSrcSlides = src1.getSlides().size() + src2.getSlides().size() + src3.getSlides().size();

        Presentation resultPres = new Presentation(resultBlob.getStream());
        int countMergedSlides = resultPres.getSlides().size();
        assertEquals(countSrcSlides, countMergedSlides);

        JSONObject resultInfo = pptUtils.getProperties(resultBlob);

        // For each presentation the master slides have been copied, even
        // if all presentations have the same.
        // So we must have 3 themes.
        JSONArray resultMasters = resultInfo.getJSONArray("MasterSlides");
        assertEquals(3, resultMasters.length());

        // Now, check we do have the specific theme stored in merge3.pptx
        JSONObject merge3Info = pptUtils.getProperties(blob3);
        JSONArray merge3Masters = merge3Info.getJSONArray("MasterSlides");
        assertEquals(1, merge3Masters.length());
        String merge3Theme = merge3Masters.getJSONObject(0).getString("Name");
        int merge3CountLayouts = merge3Masters.getJSONObject(0).getJSONArray("Layouts").length();

        boolean found = false;
        for (int i = 0; i < resultMasters.length(); i++) {
            JSONObject masterInfo = resultMasters.getJSONObject(i);
            if (masterInfo.getString("Name").equals(merge3Theme)) {
                found = true;
                assertEquals(merge3CountLayouts, masterInfo.getJSONArray("Layouts").length());
                break;
            }
        }
        assertTrue("Theme form merge3 deck not found in the result", found);

    }

    @Test
    public void shouldMergePresentationsWithReuseMasterSlides() throws Exception {

        BlobList blobs = new BlobList();

        // (See Class comments)
        File fileMerge1 = FileUtils.getResourceFileFromContext("files/merge1.pptx");
        Blob blob1 = new FileBlob(fileMerge1);
        blobs.add(blob1);
        File fileMerge2 = FileUtils.getResourceFileFromContext("files/merge2.pptx");
        Blob blob2 = new FileBlob(fileMerge2);
        blobs.add(blob2);
        File fileMerge3 = FileUtils.getResourceFileFromContext("files/merge3.pptx");
        Blob blob3 = new FileBlob(fileMerge3);
        blobs.add(blob3);

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        Blob resultBlob = pptUtils.merge(blobs, true, null);

        assertNotNull(resultBlob);

        // TestUtils.saveBlobOnDesktop(resultBlob, "test-ppt-utils");

        // We passed null as fileName => the code should provide the default name
        assertEquals("merged.pptx", resultBlob.getFilename());

        // Check slides
        Presentation src1 = new Presentation(blob1.getStream());
        Presentation src2 = new Presentation(blob2.getStream());
        Presentation src3 = new Presentation(blob3.getStream());
        int countSrcSlides = src1.getSlides().size() + src2.getSlides().size() + src3.getSlides().size();

        Presentation resultPres = new Presentation(resultBlob.getStream());
        int countMergedSlides = resultPres.getSlides().size();
        assertEquals(countSrcSlides, countMergedSlides);

        // Perform the assertions from the properties.
        JSONObject resultInfo = pptUtils.getProperties(resultBlob);

        // We reused the master slides. So we must have only 2 themes (See Class comments)
        JSONArray resultMasters = resultInfo.getJSONArray("MasterSlides");
        assertEquals(2, resultMasters.length());

        // Now, check we do have the specific theme stored in merge3.pptx
        // See merge3.pptx, View > Master > Slide Master
        String merge3Theme = "Test-Theme";
        int merge3CountLayouts = 2;

        boolean found = false;
        for (int i = 0; i < resultMasters.length(); i++) {
            JSONObject masterInfo = resultMasters.getJSONObject(i);
            if (masterInfo.getString("Name").equals(merge3Theme)) {
                found = true;
                assertEquals(merge3CountLayouts, masterInfo.getJSONArray("Layouts").length());
                break;
            }
        }
        assertTrue("Theme form merge3 deck not found in the result", found);

    }

    @Test
    public void testGetThumbnailsWithDefaultValues() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();

        BlobList blobs = pptUtils.getThumbnails(testFileBlob, 0, null, false);

        assertTrue(blobs.size() > 0);

        // For quick tests on your Mac :-)
        // for (Blob b : blobs) {
        // TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        // }

        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {
            assertEquals(fullPres.getSlides().size(), blobs.size());
            Dimension pgsize = fullPres.getPageSize();
            int w = pgsize.width;
            int h = pgsize.height;

            for (Blob b : blobs) {
                assertEquals(b.getMimeType(), "image/png");

                ImageInfo info = imagingService.getImageInfo(b);
                assertEquals(w, info.getWidth());
                assertEquals(h, info.getHeight());
            }
        }
    }

    @Test
    public void testGetThumbnailsIgnoreHidden() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();

        BlobList blobs = pptUtils.getThumbnails(testFileBlob, 0, null, true);

        assertTrue(blobs.size() > 0);

        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {
            // We have one hidden slide in this test file
            assertEquals(fullPres.getSlides().size() - 1, blobs.size());
        }

    }

    @Test
    public void testGetThumbnailsAsJpeg() throws Exception {

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();

        BlobList blobs = pptUtils.getThumbnails(testFileBlob, 0, "jpeg", false);

        assertTrue(blobs.size() > 0);

        // For quick tests on your Mac :-)
        // for (Blob b : blobs) {
        // TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        // }

        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {
            assertEquals(fullPres.getSlides().size(), blobs.size());
            Dimension pgsize = fullPres.getPageSize();
            int w = pgsize.width;
            int h = pgsize.height;

            for (Blob b : blobs) {
                assertEquals(b.getMimeType(), "image/jpeg");

                ImageInfo info = imagingService.getImageInfo(b);
                assertEquals(w, info.getWidth());
                assertEquals(h, info.getHeight());
            }
        }
    }

    @Test
    public void tesGetSlide() throws Exception {

        int SLIDE_NUMBER = 4;

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        Blob result = pptUtils.getSlide(testFileBlob, SLIDE_NUMBER);

        assertNotNull(result);
        // First slide is numbered 1, not zero (see PowerPointUtils interface)
        assertTrue(result.getFilename().endsWith("-" + (SLIDE_NUMBER + 1) + ".pptx"));

        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {

            XSLFSlide original = fullPres.getSlides().get(SLIDE_NUMBER);
            try (XMLSlideShow pres = new XMLSlideShow(result.getStream())) {

                assertEquals(1, pres.getSlides().size());

                XSLFSlide slide = pres.getSlides().get(0);
                assertTrue(TestUtils.slidesLookTheSame(original, slide));
            }
        }
    }

    @Test
    public void testGetOneThumbnailWithDefaultParameters() throws Exception {

        int SLIDE_NUMBER = 4;

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        Blob result = pptUtils.getThumbnail(testFileBlob, SLIDE_NUMBER, 0, null);

        assertNotNull(result);
        // Default is png
        assertEquals(result.getMimeType(), "image/png");
        // First slide is numbered 1, not zero (see PowerPointUtils interface)
        assertEquals("Slide " + (SLIDE_NUMBER + 1) + ".png", result.getFilename());

        ImageInfo info = imagingService.getImageInfo(result);
        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {

            Dimension pgsize = fullPres.getPageSize();
            int w = pgsize.width;
            int h = pgsize.height;

            assertEquals(w, info.getWidth());
            assertEquals(h, info.getHeight());
        }

    }

    @Test
    public void shouldGetOneThumbnailAsSmallJpeg() throws Exception {

        int SLIDE_NUMBER = 4;

        Blob testFileBlob = TestUtils.getMainTestPresentationTest();

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        Blob result = pptUtils.getThumbnail(testFileBlob, SLIDE_NUMBER, 200, "jpeg");

        assertNotNull(result);
        assertEquals(result.getMimeType(), "image/jpeg");
        // First slide is numbered 1, not zero (see PowerPointUtils interface)
        assertEquals("Slide " + (SLIDE_NUMBER + 1) + ".jpg", result.getFilename());

        ImageInfo info = imagingService.getImageInfo(result);
        assertEquals(200, info.getWidth());

    }
    
    // Cannot test Aspose with evaluation version
    @Test
    @Ignore
    public void shouldReplaceText() throws Exception {
        
        File testFile = FileUtils.getResourceFileFromContext("files/template.pptx");
        Blob template = new FileBlob(testFile);
        // Get the original template for comparison, later
        SimpleBlobHolder blobHolder = new SimpleBlobHolder(template);
        BlobHolder resultBlob = conversionService.convert("any2text", blobHolder, null);
        String templateText = new String(resultBlob.getBlob().getByteArray(), "UTF-8");
        
        DocumentModel doc = session.createDocumentModel("/", "testfile", "File");
        doc.setPropertyValue("dc:title", "The Title");
        doc.setPropertyValue("dc:description", "The Description\nWith one line");
        
        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        Blob result = pptUtils.renderWithTemplate(doc, template, null);
        assertNotNull(result);
        
        blobHolder = new SimpleBlobHolder(result);
        resultBlob = conversionService.convert("any2text", blobHolder, null);
        String finalText = new String(resultBlob.getBlob().getByteArray(), "UTF-8");
        assertNotNull(finalText);
        assertNotEquals(templateText, finalText);
        assertTrue(finalText.indexOf("The Title") > -1);
        assertTrue(finalText.indexOf("The Description\nWith one line") > -1);
        assertTrue(finalText.indexOf("Administrator") > -1);
    }

    @Test
    public void quickTest() throws Exception {
        // . . .
    }

}
