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

import org.apache.commons.lang3.StringUtils;
import org.json.JSONObject;
import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.ecm.platform.mimetype.MimetypeDetectionException;
import org.nuxeo.ecm.platform.mimetype.MimetypeNotFoundException;
import org.nuxeo.ecm.platform.mimetype.interfaces.MimetypeRegistry;
import org.nuxeo.ecm.platform.mimetype.service.MimetypeRegistryService;
import org.nuxeo.runtime.api.Framework;

public interface PowerPointUtils {

    /**
     * Returns a JSONObject with the presentation properties.
     * TODO: Use an interface, maybe, to harmonize the values when we add different providers (Apache POI, Aspose, ...)
     * 
     * @param blob
     * @return
     * @since 10.10
     */
    JSONObject getProperties(Blob blob);

    /**
     * Returns a list of blob, one/slide in the input presentation. If the input presentation is null or is not a
     * PowerPoint file, returns and empty list (not null)
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
     * 
     * @param input, the document containing a PowerPoint presentation
     * @param xpath, the field storing the presentation. Optional, "file:content" by default
     * @return the list of blob, one/slide.
     * @since 10.10
     */
    BlobList splitPresentation(DocumentModel input, String xpath) throws IOException;
    
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

}
