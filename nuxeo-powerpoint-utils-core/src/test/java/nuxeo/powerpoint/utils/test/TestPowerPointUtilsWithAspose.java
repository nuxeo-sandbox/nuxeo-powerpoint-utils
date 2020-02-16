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

import static org.junit.Assert.*;

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
import org.nuxeo.runtime.test.runner.Deploy;
import org.nuxeo.runtime.test.runner.Features;
import org.nuxeo.runtime.test.runner.FeaturesRunner;

import nuxeo.powerpoint.utils.apachepoi.PowerPointUtilsWithApachePOI;
import nuxeo.powerpoint.utils.aspose.PowerPointUtilsWithAspose;

/**
 * @since 10.10
 */
@RunWith(FeaturesRunner.class)
@Features(AutomationFeature.class)
@RepositoryConfig(init = DefaultRepositoryInit.class, cleanup = Granularity.METHOD)
@Deploy("nuxeo.powerpoint.utils-core")
public class TestPowerPointUtilsWithAspose {

    public static final String BIG_PRESENTATION = "files/2020-Nuxeo-Overview-abstract.pptx";

    @Inject
    protected CoreSession session;

    @Test
    public void shouldSplitABlobPresentation() throws Exception {

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        BlobList blobs = pptUtils.splitPresentation(testFileBlob);

        assertNotNull(blobs);

        // For quick tests on your Mac :-)
        // for (Blob b : blobs) {
        // TestUtils.saveBlobOnDesktop(b, "test-ppt-utils");
        // }

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
                        assertTrue(countSlides == 1 || countSlides == 2);

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

        File testFile = FileUtils.getResourceFileFromContext(BIG_PRESENTATION);
        assertNotNull(testFile);
        Blob testFileBlob = new FileBlob(testFile);
        assertNotNull(testFileBlob);

        testFileBlob.setMimeType("application/vnd.openxmlformats-officedocument.presentationml.presentation");

        PowerPointUtilsWithAspose pptUtils = new PowerPointUtilsWithAspose();
        JSONObject result = pptUtils.getProperties(testFileBlob);

        System.out.println("\n" + result.toString(2));

        // See, in PowerPoint, File > Properties of the test file.
        assertEquals("Nuxeo Unit Testing", result.get("Creator"));
        assertEquals("Nuxeo", result.get("Company"));
        assertEquals("Widescreen", result.get("PresentationFormat"));
        
        // TODO
        /*
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
        */
    }

}
