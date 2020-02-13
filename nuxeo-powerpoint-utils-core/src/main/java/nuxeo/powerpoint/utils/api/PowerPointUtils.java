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

import org.nuxeo.ecm.automation.core.util.BlobList;
import org.nuxeo.ecm.core.api.Blob;

public interface PowerPointUtils {

    /**
     * Returns a list of blob, one/slide in the input presentation. If the input presentation is null or is not a
     * PowerPoint file, returns and empty list (not null)
     * 
     * @param input, the blob containing the presentation to split
     * @return a list of blobs, one/slide. Empty list is input is null or not a presentation
     * @since TODO
     */
    BlobList splitPresentation(Blob blob) throws IOException;

    Blob mergeSlides(BlobList slides);

}
