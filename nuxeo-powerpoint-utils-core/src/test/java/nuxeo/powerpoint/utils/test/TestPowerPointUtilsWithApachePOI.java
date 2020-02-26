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
import static org.junit.Assert.assertFalse;
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
import org.junit.Test;
import org.junit.runner.RunWith;
import org.nuxeo.common.utils.FileUtils;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.automation.test.AutomationFeature;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.ecm.core.test.DefaultRepositoryInit;
import org.nuxeo.ecm.core.test.annotations.Granularity;
import org.nuxeo.ecm.core.test.annotations.RepositoryConfig;
import org.nuxeo.ecm.platform.picture.api.ImageInfo;
import org.nuxeo.ecm.platform.picture.api.ImagingService;
import org.nuxeo.runtime.test.runner.Deploy;
import org.nuxeo.runtime.test.runner.Features;
import org.nuxeo.runtime.test.runner.FeaturesRunner;

import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;

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
@Deploy({ "org.nuxeo.ecm.platform.picture.api", "org.nuxeo.ecm.platform.picture.core",
        "org.nuxeo.ecm.platform.picture.convert", "org.nuxeo.ecm.platform.tag", "nuxeo.powerpoint.utils-core" })
public class TestPowerPointUtilsWithApachePOI {

    public static final String BIG_PRESENTATION = "files/2020-Nuxeo-Overview-abstract.pptx";

    @Inject
    protected CoreSession session;

    @Inject
    protected ImagingService imagingService;

    @Test
    public void shouldSplitABlobPresentation() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
        BlobList blobs = pptUtils.splitPresentation(testFileBlob);

        assertNotNull(blobs);

        // For quick tests on your Mac :-)
        // for (Blob b : blobs) {
        // TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        // }

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

    // As of "today", merging works only with Aspose.
    @Test
    public void shouldFailMergingSlides() throws Exception {
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

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
        try {
            @SuppressWarnings("unused")
            Blob result = pptUtils.merge(blobs, false, null);
            assertFalse("Merging with Apache OI should fail. If working, update documentaiton and this test.", true);
        } catch (Exception e) {

        }

    }

    @Test
    public void tesGetProperties() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();
        JSONObject result = pptUtils.getProperties(testFileBlob);

        // See, in PowerPoint, File > Properties of the test file.
        assertEquals("Nuxeo Unit Testing", result.get("Creator"));
        assertEquals("Nuxeo", result.get("Company"));
        assertEquals("Widescreen", result.get("PresentationFormat"));
        assertEquals(11, result.get("CountSlides"));
        assertEquals(1, result.get("CountHiddenSlides"));

        JSONArray arr = result.getJSONArray("MasterSlides");
        assertEquals(2, arr.length());
        // First one is "Office Theme"
        JSONObject theme = arr.getJSONObject(0);
        // getJSONObject does not return null is there is no value, it throws an exception
        assertEquals("Office Theme", theme.get("Name"));
        // Could also check the layouts...

        // Could also check info on every slides...

        // System.out.println("\n" + result.toString(2));
    }

    @Test
    public void testGetThumbnailsWithDefaultValues() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();

        BlobList blobs = pptUtils.getThumbnails(testFileBlob, 0, null, false);

        assertTrue(blobs.size() > 0);

        // For quick tests on your Mac :-)
        //for (Blob b : blobs) {
        //    TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        //}

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

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();

        BlobList blobs = pptUtils.getThumbnails(testFileBlob, 0, null, true);

        assertTrue(blobs.size() > 0);

        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {
            // We have one hidden slide in this test file
            assertEquals(fullPres.getSlides().size() - 1, blobs.size());
        }
    }

    @Test
    public void testGetThumbnailsSmallerThumbnails() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();

        BlobList blobs = pptUtils.getThumbnails(testFileBlob, 200, null, false);

        assertTrue(blobs.size() > 0);

        // For quick tests on your Mac :-)
        // for (Blob b : blobs) {
        // TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        // }

        try (XMLSlideShow fullPres = new XMLSlideShow(testFileBlob.getStream())) {
            assertEquals(fullPres.getSlides().size(), blobs.size());

            for (Blob b : blobs) {
                assertEquals(b.getMimeType(), "image/png");

                ImageInfo info = imagingService.getImageInfo(b);
                assertEquals(200, info.getWidth());
            }
        }
    }

    @Test
    public void testGetThumbnailsAsJpeg() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        PowerPointUtilsWithApachePOI pptUtils = new PowerPointUtilsWithApachePOI();

        BlobList blobs = pptUtils.getThumbnails(testFileBlob, 0, "jpg", false);

        assertTrue(blobs.size() > 0);

        // For quick tests on your Mac :-)
        //for (Blob b : blobs) {
        //    TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        //}

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

}
