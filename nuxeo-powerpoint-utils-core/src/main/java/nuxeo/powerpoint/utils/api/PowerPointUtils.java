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
package nuxeo.powerpoint.utils.api;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.commons.lang3.StringUtils;
import org.json.JSONObject;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.DocumentModelList;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.ecm.platform.mimetype.MimetypeDetectionException;
import org.nuxeo.ecm.platform.mimetype.MimetypeNotFoundException;
import org.nuxeo.ecm.platform.mimetype.interfaces.MimetypeRegistry;
import org.nuxeo.ecm.platform.mimetype.service.MimetypeRegistryService;
import org.nuxeo.runtime.api.Framework;

public interface PowerPointUtils {

    public static final String PPTX_MIMETYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

    // Use this DateFormat to format the dates in <code>JSONObject getProperties(Blob blob)</code>
    // For example: <code>obj.put("Created", DATE_FORMAT.format(yourPre.getADate()));</code>
    public static final DateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS");

    /**
     * Returns a JSONObject with the presentation properties.
     * TODO: Use an interface, maybe, to harmonize the values when we add different providers (Apache POI, Aspose, ...)
     * 
     * @param blob, the PowerPoint presentation
     * @return a JSON object with a list of properties
     * @since 10.10
     */
    JSONObject getProperties(Blob blob);

    /**
     * Returns a list of blob, one/slide in the input presentation. If the input presentation is null or is not a
     * PowerPoint file, returns and empty list (not null).
     * The name of each blob will be "{original-filename-}{slideNumber + 1}.pptx", starting at 1.
     * 
     * @param input, the blob containing the presentation to split
     * @return a list of blobs, one/slide. Empty list if input is null or not a presentation
     * @since 10.10
     */
    BlobList splitPresentation(Blob blob) throws IOException;

    /**
     * Returns a list of blobs, one/slide after splitting the presentation contained in the input document in the xpath
     * field (if null or empty, default to "file:content"). Returns an empty list in the blob at xpath is null, or is
     * not a presentation.
     * The name of each blob will be "{original-filename-}{slideNumber + 1}.pptx", starting at 1.
     * 
     * @param input, the document containing a PowerPoint presentation
     * @param xpath, the field storing the presentation. Optional, "file:content" by default
     * @return the list of blob, one/slide.
     * @since 10.10
     */
    BlobList splitPresentation(DocumentModel input, String xpath) throws IOException;

    /**
     * Merge all presentations to a single one, in the received order.
     * If <code>reuseMasters</code> is <code>false</code>, the slide's layouts and styles are preserved,
     * which means the master slides will be duplicated. If <code>true</code>, the code will reuse a master
     * slide of same theme and same layout already existing in the merged presentation being builtt.
     * If <code>fileName</code> is null or empty, the file name is set to "merged.pptx"
     * Always create a.pptx blob. Adds ".pptx" to fileName if it does not end with .pptx.
     * If any of these condition applies for any blob in <code>blobs</code> is ignored (no conversion applies).
     * Also, when a blob in <code>blobs</code> has zero slide, it is ignored.
     * If <code>blobs</code> is null or empty, null is returned.
     * 
     * @param blobs, the PowerPoint presentations
     * @param reuseMasters
     * @param fileName
     * @return the presentation merging all the input blobs
     * @since 10.10
     */
    Blob merge(BlobList blobs, boolean reuseMasters, String fileName);

    /**
     * Extract all the blobs stored in each documents at <code>xpath</xpath> (default to "file:content") and
     * just calls <code>Blob merge(BlobList blobs, boolean reuseMasters, String fileName);</code>
     * 
     * @param docs, holding the PowerPoint presentations
     * @param xpath, the field storing the presentation. Optional, "file:content" by default
     * @param reuseMasters
     * @param fileName
     * @return
     * @since 10.10
     */
    Blob merge(DocumentModelList docs, String xpath, boolean reuseMasters, String fileName);

    /**
     * Return a presentation of one slide. Master slides are added to the slide.
     * slideNumber is zero-based, but the file name will be...
     * The name of the file will be "{original-filename-}{slideNumber + 1}.pptx
     * ... so it is not necessary to re-process the titles for end users
     * 
     * @param blob, the PowerPoint presentation
     * @param slideNumber, zero-based
     * @return a presentation containing only the slide.
     * @throws IOException
     * @since 10.10
     */
    Blob getSlide(Blob blob, int slideNumber) throws IOException;

    /**
     * Return a presentation of one slide. Source presentation is read in the input document in the xpath field. If
     * xpath is empty or null, "file:content" is used. Master slides are added to the slide.
     * slideNumber is zero-based, but the file name will be...
     * The name of the file will be "{original-filename-}{slideNumber + 1}.pptx
     * ... so it is not necessary to re-process the titles for end users
     * 
     * @param input
     * @param xpath
     * @param slideNumber
     * @return
     * @throws IOException
     * @since 10.10
     */
    Blob getSlide(DocumentModel input, String xpath, int slideNumber) throws IOException;

    /**
     * Returns a list of images, one thumbnail/slide contained in blob presentation, with options:
     * - A maximum width. If this width is lower than the presentation width, then the height will also be reduced
     * accordingly. Any value <= 0 means "original SlideDeck size"
     * - format can be "jpg", "jpeg" or "png". Any other format thows an exception. If not empty or null, use "png"
     * - onlyVisible: if true, hidden slides will be ignored, no thumbnail will be calculated
     * Each blob's file name is "Slide {slideNumber}.pptx". Number starts at 1 (for better end user experience)
     * 
     * @param blob, the presentation
     * @param maxWidth
     * @param format, "jpg", "jpeg" or "png" only
     * @param onlyVisible, if true, no thumbnail will be calculated for hidden slides
     * @return a list of images in the desired format and size
     * @throws IOException
     * @since 10.10
     */
    BlobList getThumbnails(Blob blob, int maxWidth, String format, boolean onlyVisible) throws IOException;

    /**
     * Returns a list of images, one thumbnail/slide contained in the presentation in doc, in the xpath
     * field (if null or empty, default to "file:content")with options:
     * - A maximum width. If this width is lower than the presentation width, then the height will also be reduced
     * accordingly. Any value <= 0 means "original SlideDeck size"
     * - format can be "jpg", "jpeg" or "png". Any other format thows an exception. If not empty or null, use "png"
     * - onlyVisible: if true, hidden slides will be ignored, no thumbnail will be calculated
     * Each blob's file name is "Slide {slideNumber}.pptx". Number starts at 1 (for better end user experience)
     * 
     * @param doc, document holging the presentation
     * @param xpath, the field to use (default to "file:content")
     * @param maxWidth
     * @param format, "jpg", "jpeg" ot "png" only
     * @param onlyVisible, if true, no thumbnail will be calculated for hidden slides
     * @return a list of images in the desired format and size
     * @throws IOException
     * @since 10.10
     */
    BlobList getThumbnails(DocumentModel doc, String xpath, int maxWidth, String format, boolean onlyVisible)
            throws IOException;

    /**
     * Return the thumbnail of one slide inside the presentation stored in blob, with these parameters:
     * - slideNumber: a zero-based value. But the file name will be... "Slide {slideNumber + 1}.pptx"
     * (number starts at 1 for better end user experience)
     * - maxWidth: A maximum width. If this width is lower than the presentation width, then the height will also be
     * reduced accordingly. Any value <= 0 means "original SlideDeck size"
     * - format can be "jpg", "jpeg" or "png". Any other format thows an exception. If not empty or null, use "png"
     * 
     * @param blob, the PowerPoint presentation
     * @param maxWidth
     * @param format, "jpg", "jpeg" or "png" only
     * @return the thumbnail
     * @since 10.10
     */
    Blob getThumbnail(Blob blob, int slideNumber, int maxWidth, String format) throws IOException;

    /**
     * Return the thumbnail of one slide inside the presentation stored in the document at xpath (if null or empty,
     * defaults to "file:content"), with these parameters:
     * - slideNumber: a zero-based value. But the file name will be... "Slide {slideNumber + 1}.pptx"
     * (number starts at 1 for better end user experience)
     * - maxWidth: A maximum width. If this width is lower than the presentation width, then the height will also be
     * reduced accordingly. Any value <= 0 means "original SlideDeck size"
     * - format can be "jpg", "jpeg" or "png". Any other format thows an exception. If not empty or null, use "png"
     * Fetches the
     * 
     * @param blob, the PowerPoint presentation
     * @param maxWidth
     * @param format, "jpg", "jpeg" or "png" only
     * @return the thumbnail
     * @since 10.10
     */
    Blob getThumbnail(DocumentModel doc, String xpath, int slideNumber, int maxWidth, String format) throws IOException;

    /**
     * Receives a document and a pptx template. processes the template, replaces freemarker tag with the values, return
     * a new pptx with the result.
     * 
     * Check each implementation for limitation. As of first implementation:
     *   - Apache POI:
     *     • Only ${...} expressions are supported (so no loops, no <#if, etc.)
     *     • The expression must be embedded in lines.
     *        So, for example, this...
     *                  ----------------------
     *                 |  Hello               |
     *                 |  ${doc["dc:title"]}  |
     *                  ----------------------
     *        ... will fail, while this one works well:
     *                  ----------------------------
     *                 |  Hello ${doc["dc:title"]}  |
     *                  ----------------------------
     *                 Or:
     *                  ----------------------
     *                 |  Hello               |
     *                  ----------------------
     *                  ----------------------
     *                 |  ${doc["dc:title"]}  |
     *                  ----------------------
     * 
     *   - Aspose: Impossible to test without a license, because it fails
     *             to extract more than a few characters.
     * @param doc
     * @param template
     * @param newFileName
     * @return
     * @throws Exception
     * @since 2021.35
     */
    public Blob renderWithTemplate(DocumentModel doc, Blob template, String newFileName) throws Exception;

    /**
     * Helper utility getting the mime-type of a blob
     * 
     * @param blob
     * @since 10.10
     */
    public static String getBlobMimeType(Blob blob) {

        if (blob == null) {
            throw new NullPointerException();
        }

        String mimeType = blob.getMimeType();
        if (StringUtils.isNotBlank(mimeType)) {
            return mimeType;
        }

        MimetypeRegistryService service = (MimetypeRegistryService) Framework.getService(MimetypeRegistry.class);
        try {
            mimeType = service.getMimetypeFromBlob(blob);
        } catch (MimetypeNotFoundException | MimetypeDetectionException e1) {
            try {
                mimeType = service.getMimetypeFromFile(blob.getFile());
            } catch (MimetypeNotFoundException | MimetypeDetectionException e2) {
                throw new NuxeoException("Cannot get a Mime Type from the blob or the file", e2);
            }
        }

        return mimeType;
    }

    /**
     * Utility to ensure all providers use the name as stated in the interface.
     * A merged presentation must ends with .pptx. If it is not the case, we add the ".pptx" suffix.
     * A default name ("merged.pptx") is provided if fioeName is null or empty.
     * 
     * @param fileName
     * @return The fileName with the correct name/extension
     * @since 10.10
     */
    public static String checkMergedFileName(String fileName) {

        if (StringUtils.isBlank(fileName)) {
            return "merged.pptx";
        }

        if (!fileName.toLowerCase().endsWith(".pptx")) {
            return fileName + ".pptx";
        }

        return fileName;
    }

    /**
     * Centralizing code repeated in a lot of places, to get "file:content" if no xpath is provided
     * 
     * @param doc
     * @param xpath
     * @return the blob at xpath, or at "file:content"
     * @since 10.10
     */
    public static Blob getBlob(DocumentModel doc, String xpath) {

        if (StringUtils.isBlank(xpath)) {
            xpath = "file:content";
        }
        return (Blob) doc.getPropertyValue(xpath);
    }

}
