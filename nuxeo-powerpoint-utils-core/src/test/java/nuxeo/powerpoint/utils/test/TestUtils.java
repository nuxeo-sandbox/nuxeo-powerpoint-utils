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

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFComments;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.nuxeo.ecm.core.api.Blob;

/**
 * @since 10.10
 */
public class TestUtils {

    /*
     * This one is for local quick test with human checking :-). Requires inFolderName
     * to exist on your Desktop
     */
    public static void saveBlobOnDesktop(Blob inBlob, String inFolderName) throws IOException {
        File destFile = new File(System.getProperty("user.home"),
                "Desktop/" + inFolderName + "/" + inBlob.getFilename());
        inBlob.transferTo(destFile);
    }

    /**
     * Return true if both slides look equal (else, return false)
     * 
     * @param s1
     * @param s2
     * @return true if both slides look equal (else, return false)
     * @since 10.10
     */
    public static boolean slidesLookTheSame(XSLFSlide s1, XSLFSlide s2) {
        
        return slideToString(s1).equals(slideToString(s2));
    }
    
    /**
     * Return a String description of some properties of the slide.
     * 
     * @param slide
     * @return String description of some properties of the slide
     * @since 10.10
     */
    public static String slideToString(XSLFSlide slide) {
        
        ArrayList<String> values = new ArrayList<String>();
        
        values.add("Title: " + slide.getTitle());
                
        XSLFComments comments = slide.getComments();
        if(comments != null) {
            values.add("Comments: " + comments.getNumberOfComments());
        }
        
        values.add("Layout name: " + slide.getSlideLayout().getName());
        values.add("Bg fillcolor: " + slide.getBackground().getFillColor());
        values.add("Relastions: " + slide.getRelations().size());
                
        return String.join(", ", values);
    }
}
