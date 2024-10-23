# nuxeo-powerpoint-utils

This plugin for [Nuxeo Platform](http://www.nuxeo.com) allows for handling PowerPoint sides: Extract information, split, render with template, ...

The plugin uses [Apache POI](https://poi.apache.org).

> [!IMPORTANT]
> * Versions 2.0.n are for LTS 2021 and can use [Aspose Slides](https://products.aspose.com/slides) instead of Apache POI, _but_ some features (like merging slides) are only available using Aspose, which requires a valid license. Without such valid license key, Aspose can only be used for testing and slides created with the tool are [watermarked](https://docs.aspose.com/display/slidesjava/Licensing). (See below for more details)
> * Starting with Nuxeo LTS 2023 (Nuxeo PowerPoint utils version 2O23.n), the plugin does not use aspose anymore (no time to debug issues when upgrading, in the context of this sandbox plugin), and calls to operations with Aspose throw and error, "Version 2023.n of Nuxeo PowerPoint Utilities does not support Aspose"

# Table of Content
- [Usage](#usage)
  * [Conversion.PowerPointGetProperties](#conversionpowerpointgetproperties)
  * [Conversion.PowerPointMerge](#conversionpowerpointmerge)
  * [Conversion.PowerPointSplit](#conversionpowerpointsplit)
  * [Conversion.PowerPointGetSlide](#conversionpowerpointgetslide)
  * [Conversion.PowerPointGetThumbnails](#conversionpowerpointgetthumbnails)
  * [Conversion.PowerPointGetOneThumbnail](#conversionpowerpointgetonethumbnail)
  * [Conversion.RenderDocumentWithPowerPointTemplate](#conversionrenderdocumentwithpowerpointtemplate)
  * [Conversion.SetAsposeSlidesLicense](#conversionsetasposeslideslicense)
- [Apache POI vs Aspose](#apache-poi-vs-aspose)
- [Example of Properties Output](#example-of-properties-output)
- [Build](#build)
- [Support](#support)
- [Licensing](#licensing)
- [About Nuxeo](#about-nuxeo)

# Usage
The plugin provides utilities for extracting info, splitting and merging PowerPoint presentations (for Java developer, see the `PowerPointUtils` interface). These utilities can be used via operations described here.

#### Conversion.PowerPointGetProperties
* Label: `PowerPoint: Get Properties`
* Input: `Blob` or `Document`
* Output: `String`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a document. `xpath` is the field to use
    * Default value is `"file:content"`
  * `useAspose`
    * Boolean, optional (default: `false`)
    * When using Aspose, more information can be returned, like the list of fonts used in the presentation.
* Return a JSON string containing the properties. See below "Example of Properties Output"

#### Conversion.PowerPointMerge

**WARNING: This operation uses Aspose only, it is not possible to use Apache POI for this purpose**

* Label: `PowerPoint: Merge Presentations`
* Input: `Blobs` or `Documents`
* Output: `Blob`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a `Documents`. `xpath` is the field to use
    * Default value is `"file:content"`
  * `fileName`
    * String, optional
    * The name of the resulting file. if it does not end with ".pptx", it is added.
    * Default: "merged.pptx"
  * `reuseMasters`
    * Boolean, optional
    * If `false`, all the master slides of the source presentations are added to the final, merged ones. This means that if some input presentations use the same masters, they will be duplicated in the resulting, merged presentation.
    * When `true`, the operation will transfer a copy of the original master slides only if they don't already exist in the merged presentation.
    * This is based on the combination _theme name + layout name_.
* Returns a `Blob`, the presentation merging all the input ones. It is always a `pptx` presentation.

#### Conversion.PowerPointSplit

Split the input presentation and returns a list of blobs, one per slide. Each slide also contains a copy of the original master slides (the theme) used. For each blob, the file name is: `{original presentation name}-{slideNumberStartAt1}.pptx` (starts at 1, not zero, so there is less confusion for an end user)

**Warning #1**: If the master slides of the input presentation are "bigs" (contain HiRes images, videos, ...), then each slide will be big too. For example, if the size of all the master slides is 40MB and there are 100 slides, each slide will be at least 40MB, which is normal, they also contain the HiRes images, the videos, etc.

**WARNING #2***: The operation can take several seconds to complete. If the presentation to split contains dozens of complex and heavy slides (images, videos, ...), it can take dozens of seconds. By default it does not run an asynchronous worker. We recommend launching the operation asynchronously (and maybe add a mail notification once the split is done).

* Label: `PowerPoint: Split Presentation`
* Input: `Blob` or `Document`
* Output: `BlobList`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a `Document`. `xpath` is the field to use
    * Default value is `"file:content"`
  * `useAspose`
    * boolean, optional (default: `false`)
    * If `false` (default value), the code will make use of Apache POI to split the presentation.
    * **WARNING** On this case, splitting the presentation can be slow. For big presentation (dozens of complex slides), we recommend running it asynchronously if it was launched by a user in the UI. With Nuxeo Automation, it is possible to handle the business logic and then send a mail notification once the split is done.
    * If `true`, the operation will use Aspose to split the slides. This is done very quickly. This requires a valid Aspose license
* Returns a `BlobList`, list of `Blobs`. Each blob is a side of the input presentation. It also contains a copy of the master slides.

#### Conversion.PowerPointGetSlide
Return a `Blob`, single slide presentation, copy of the slide passed in the `slideNumber` parameter. The master slides are always copied to the returned presentation.

The result blob's file name is `{original presentation name}-{slideNumberStartAt1}.pptx"`. So, even if `slideNumber` is zero-based, the file name is 1-based. This is done to avoid having users wondering why they requested slide 4 and got it, but named "my Presentation-3.pptx"

* Label: `PowerPoint: Get Slide`
* Input: `Blob` or `Document`
* Output: `Blob`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a `Document`. `xpath` is the field to use
    * Default value is `"file:content"`
  * `slideNumber`
    * Integer, _required_
    * The number of the slide to extract. 0-based (value must be between 0 and (number of slides - 1)
  * `useAspose`
    * boolean, optional (default: `false`)
    * If `false` (default value), the code will make use of Apache POI, else it uses Aspose
    * Aspose generates, usually, smaller slides with the same quality.
* Returns a `Blob`, a powerpoint presentation with the single slide

#### Conversion.PowerPointGetThumbnails

Return a `BlobList` of thumbnails, one/slide, as PNG of JPEG, in the original slide dimensions or with a scale factor. It is possible to return only the visible slides.

* Label: `PowerPoint: Get Thumbnails`
* Input: `Blob` or `Document`
* Output: `BlobList`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a `Document`. `xpath` is the field to use
    * Default value is `"file:content"`
  * `maxWidth`
    * integer, optional
    * Allows for returning smaller images.
    * Any value <= 0 returns the images in the original dimension
  * `onlyVisible`
    * Boolean, optional
    * If `true`, thumbnails are returned only for visible slides.
  * `format`
    * String, optional, default is "png"
    * Can be only can be "jpg", "jpeg" or "png"
  * `useAspose`
    * boolean, optional (default: `false`)
    * If `false` (default value), the code will make use of Apache POI, else it uses Aspose
    * Slides rendered with Aspose usually have a better quality.
* Returns a `BlobList` of images, one per slide, in the desired size and format. Each image will have the name `{original-file-name}-{slideNumberStartAt1}.{format}` (slide numbers in the output start at 1 to avoid confusion for an end user)

#### Conversion.PowerPointGetOneThumbnail

Return a `Blob`, thumbnail of the slide, as PNG of JPEG, in the original slide dimensions or with a scale factor.

* Label: `PowerPoint: Get a Thumbnail`
* Input: `Blob` or `Document`
* Output: `Blob`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a `Document`. `xpath` is the field to use
    * Default value is `"file:content"`
  * `slideNumber`
    * Integer, _required_
    * The number of the slide to extract. 0-based (value must be between 0 and (number of slides - 1)
  * `maxWidth`
    * integer, optional
    * Allows for returning smaller images.
    * Any value <= 0 returns the images in the original dimension
  * `format`
    * String, optional, default is "png"
    * Can be only can be "jpg" or "png"
  * `useAspose`
    * boolean, optional (default: `false`)
    * If `false` (default value), the code will make use of Apache POI, else it uses Aspose
    * Slides rendered with Aspose usually have a better quality.
* Returns a `Blob`, an image rendition of the slide, in the desired size and format. The file name is `{original-file-name}-{slideNumberStartAt1}.{format}` **WARNING** When you request slide 3 (0-based) the output will be `... -4 ...`.

#### Conversion.RenderDocumentWithPowerPointTemplate

Create a pptx from a template and the input doc. In the pptx template, add FreeMarker expressions, such as `${doc["schema:field"]}`. The operation replaces the values and returns a new blob.

⚠️ **WARNING - KNOWN LIMITATIONS** ⚠️

* **The plugin only supports replacing expressions between `${` and `}`**:
  * It does not handle loops (`<#list...`), conditions (`<#if ...`), etc.
  * Example of supported expressions:
    * `${doc["dc:title"]}`
    * `${doc["customschema:field"]}`
    * `${Fn.getPrincipal(CurrentUser.name).getFirstName()}`
    * . . .
* An expression must be set on a single ligne. Create a new text block in PowerPoint if needed. For example, this...

```
 ----------------------
|  Hello               |
|  ${doc["dc:title"]}  |
 ----------------------
```
 
... will fail, while this one works well:

```
 ----------------------------
|  Hello ${doc["dc:title"]}  |
 ----------------------------
```
Or:

```
 ----------------------
|  Hello               |
 ----------------------
 ----------------------
|  ${doc["dc:title"]}  |
 ----------------------
```

**The operation**:

* Label: `PowerPoint: Render Document with Template`
* Input: `Document`
* Output: `Blob`
* Parameters:
  * `templateBlob`
    * Blob, required
    * Blob holding a .pptx slides deck. Inside this template, add FreeMarker expressions, such as `${doc[\"schema:field\"]}`
  * `fileName`
    * Name of the file to create
    * If fileName is empty, the returned blob will have the name of the template."
        + " useAspose tells the operation to use Aspose for the rendition. Default * `useAspose`
    * boolean, optional (default: `false`)
    * If `false` (default value), the code will make use of Apache POI, else it uses Aspose
    * Slides rendered with Aspose usually have a better quality.

#### Conversion.SetAsposeSlidesLicense

* Label: `Conversion.SetAsposeSlidesLicense`
* Input: `void`
* Output: `void`
* Parameter:
  * `licensePath`, string _required_.
    * Path to the license key provided by Aspose
    * See Aspose's [documentation](https://docs.aspose.com/display/slidesjava/Licensing).
    * Make sure the license is at a path where Nuxeo can access it (meaning, for example, the "user" used to install Nuxeo can access this document)

# Apache POI vs Aspose

In both cases we would like to issue a **WARNING**: As the plugin is using these external libraries, we will not be able to easily fix bugs in these libraries. Typical examples are the operations extracting thumbnails: Images can be different than the original slides, depending on the complexity of the slide, the rendering can fail in some area (typically rendering the fonts, for example).

The most efficient in terms of rendering and splitting is Aspose, which requires a [valid license](https://docs.aspose.com/display/slidesjava/Licensing):

* **[Apache POI](https://poi.apache.org)** has a [business-friendly license](https://poi.apache.org/legal.html) (Apache License Version 2.0) but does not provide easy ways to split and merge slides. Splitting is very slow (still, working). Merging is not implemented at all, if you need to merge slides using this plugin, you _must_ use Aspose.<br/>Also, Apache POI is less efficient rendering slide(s) to thumbnail(s)

* **[Aspose Slides](https://products.aspose.com/slides)**, on the other hand, is a commercial product specialized in handling PowerPoint slides, and, as such, it provides very efficient and fast ways to split, merge etc. Ti use Aspose and avoid watermarks on slides, you must register a valid license key file using the `Conversion.SetAsposeSlidesLicense` operation

# Example of Properties Output

Calling `Conversion.GetPowerPointPresentationProperties` with `useAspose` set to `true`. The list of properties is alphabetically ordered:

```
{
  "AppVersion": "16.0000",
  "Application": "Microsoft Macintosh PowerPoint",
  "AutoCompressPictures": true,
  "Company": "Nuxeo",
  "CompatMode": false,
  "CountHiddenSlides": 1,
  "CountLines": -1,
  "CountMMClips": 0,
  "CountNotes": 10,
  "CountPages": -1,
  "CountParagraphs": 176,
  "CountSlides": 11,
  "CountTotalTime": 14140,
  "CountWords": 599,
  "Created": "2017-10-06T20:06:38.000",
  "Creator": "Nuxeo Unit Testing",
  "EmbeddedFonts": [

  ],
  "Fonts": [
    "Arial",
    "NeueHaasGroteskDisp Std Blk",
    "NeueHaasGroteskDisp Std",
    "Wingdings",
    "Noto Sans Symbols",
    "Calibri Light",
    "Calibri",
    "Open Sans Semibold",
    "Abadi MT Condensed Extra Bold",
    "Neue Haas Grotesk Display Std 9"
  ],
  "Height": 540,
  "HyperlinkBase": "",
  "Keywords": "nuxeo,api,cloud,low-code,overview,architecture,performance",
  "LastModifiedByUser": "John Doe the Second",
  "LastPrinted": "2017-09-22T21:50:07.000",
  "Manager": "",
  "MasterSlides": [
    {
      "MasterFont": "NeueHaasGroteskDisp Std Blk",
      "Layouts": [
        "Blank",
        "Thank You Slide_2",
        "Laptop & Mobile App",
        "Title right & half background img left",
        "Title & Subtitle",
        "Title big top with bg image",
        "Cover Slide",
        "Title",
        "Agenda"
      ],
      "MinorFont": "NeueHaasGroteskDisp Std",
      "Name": "Office Theme"
    },
    {
      "MasterFont": "Calibri Light",
      "Layouts": [
        "Main2",
        "Main1",
        "Main3"
      ],
      "MinorFont": "Calibri",
      "Name": "Unit Test Second Theme"
    }
  ],
  "Modified": "2020-02-14T23:09:22.000",
  "PresentationFormat": "Widescreen",
  "Revision": "457",
  "Slidesinfo": [
    {
      "Master": "Cover Slide",
      "SlideNumber": 1,
      "Title": "",
      "Theme": "Office Theme"
    },
    {
      "Master": "Agenda",
      "SlideNumber": 2,
      "Title": "",
      "Theme": "Office Theme"
    },
    {
      "Master": "Title big top with bg image",
      "SlideNumber": 3,
      "Title": "Who is Nuxeo?",
      "Theme": "Office Theme"
    },
    {
      "Master": "Laptop & Mobile App",
      "SlideNumber": 4,
      "Title": "",
      "Theme": "Office Theme"
    },
    {
      "Master": "Laptop & Mobile App",
      "SlideNumber": 5,
      "Title": "",
      "Theme": "Office Theme"
    },
    {
      "Master": "Title & Subtitle",
      "SlideNumber": 6,
      "Title": "Where We Are. ",
      "Theme": "Office Theme"
    },
    {
      "Master": "Title big top with bg image",
      "SlideNumber": 7,
      "Title": "",
      "Theme": "Office Theme"
    },
    {
      "Master": "Blank",
      "SlideNumber": 8,
      "Title": "",
      "Theme": "Office Theme"
    },
    {
      "Master": "Title",
      "SlideNumber": 9,
      "Title": "Nuxeo Platform at a Glance",
      "Theme": "Office Theme"
    },
    {
      "Master": "Title right & half background img left",
      "SlideNumber": 10,
      "Title": "Our Kitchen is Open: Come on in.",
      "Theme": "Office Theme"
    },
    {
      "Master": "Thank You Slide_2",
      "SlideNumber": 11,
      "Title": "Thibaud @ Nuxeo",
      "Theme": "Office Theme"
    }
  ],
  "Template": "nuxeo_powerpoint-template_20171006",
  "Title": "Nuxeo Overview",
  "Width": 960
}
```

# Build

git clone https://github.com/nuxeo-powerpoint-utils.git     cd nuxeo-powerpoint-utils

mvn clean install

# Support

**These features are not part of the Nuxeo Production platform, they are not supported**

These solutions are provided for inspiration and we encourage customers to use them as code samples and learning resources.

This is a moving project (no API maintenance, no deprecation process, etc.) If any of these solutions are found to be useful for the Nuxeo Platform in general, they will be integrated directly into the platform, not maintained here.

# Licensing

[Apache License, Version 2.0](http://www.apache.org/licenses/LICENSE-2.0)

# About Nuxeo

Nuxeo dramatically improves how content-based applications are built, managed and deployed, making customers more agile, innovative and successful. Nuxeo provides a next generation, enterprise ready platform for building traditional and cutting-edge content-oriented applications. Combining a powerful application development environment with SaaS-based tools and a modular architecture, the Nuxeo Platform and Products provide clear business value to some of the most recognizable brands including Verizon, Electronic Arts, Sharp, FICO, the U.S. Navy, and Boeing. Nuxeo is headquartered in New York and Paris.

More information is available at [www.nuxeo.com](http://www.nuxeo.com).  
