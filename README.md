# nuxeo-powerpoint-utils

This plugin for [Nuxeo Platform](http://www.nuxeo.com) allows for handling PowerPoint sides: Extract information, split and merge.

#### IMPORTANT
Some features are only available using [Aspose Slides](https://products.aspose.com/slides), a third-party tool which requires a valid license. Without the key, slides created with the tool are [watermarked](https://docs.aspose.com/display/slidesjava/Licensing).

(See below for more details)


# Usage
The plugin provides utilities for extracting info, splitting and merging pPowerPoint presentations (for Java developer, see the `PowerPointUtils` interface). These utilities can be used via operations described here.

#### Conversion.GetPowerPointPresentationProperties
* Label: `PowerPoint: Get Properties`
* Input: `Blob` or `Document`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a document. `xpath` is the field to use
    * Default value is `"file:content"`
  * `useAspose`
    * When using Aspose, more information can be returned, like the list of fonts used in the presentation.
* Return a JSON string containing the properties. See below "Example of Properties Output"

#### Conversion.MergePowerPoints

<div style="background-color: lemonchiffon;margin-left:50px;margin-right:50px;font-weight: bold; padding: 10px">WARNING: This operation uses Aspose, it is not possible to use Apache POI for this purpose</div>

* Label: `PowerPoint: Merge Presentations`
* Input: `Blobs` or `Documents`
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
    * When `true`, the operation will transfer copy the original master slides only if they don't already exist in the merged presentation.
    * This is based on the combination _theme name + layout name_.
* Returns a `Blob`, the presentation merging all the input ones. It always is a `pptx` presentation.

#### Conversion.SplitPresentation

Split the input presentation and returns a list of blobs, one per slide. Each slide also contains a copy of the original master slides (the theme) used.

**Warning**: If the master slides of the input presentation are "bigs" (contain Hires images, videos, ...), then each slide will be big too. For example, if the size of all the master slides is 40MB and there are 100 slides, each slide will be at least 40MB, which is normal, they also contain the HiRes images, the videos, etc.

* Label: `PowerPoint: Split Presentation`
* Input: `Blob` or `Document`
* Parameters:
  * `xpath`:
    * String, optional
    * Used only if input is a `Document`. `xpath` is the field to use
    * Default value is `"file:content"`
  * `useAspose`
    * If `false`, the code will make use of Apache POI to split the presentation.
    * **WARNING** On this case, splitting the presentation can be slow. For big presentation (dozens of complex slides), we recommend running it asynchronously if it was launched by a user in the UI. With Nuxeo Automation, it is possible to handle the business logic and then send a mail notification once the split is done.
    * If `true`, the operation will use Aspose to split the slides. This is done very quickly. This requires a valide Aspose license
* Returns a `BlobList`, list of `Blobs`. Each blob is a side of the input presentation. It also contains a copy of the master slides
 
#### Conversion.SetAsposeSlidesLicense

# Using Aspose Slides - Limitations of Apache POI

In order to merge slides or to quickly and more efficiently split them, we recommend using Aspose, which requires a [valid license](https://docs.aspose.com/display/slidesjava/Licensing):

* **[Apache POI](https://poi.apache.org)** has a [business friendly license](https://poi.apache.org/legal.html) (Apache License Version 2.0) but does not provide easy ways to split and merge slides. Splitting is very slow (still, working). Merging is not implemented at all, if you need to merge slides using this plugin, you _must_ use Aspose.


* **[Aspose Slides](https://products.aspose.com/slides)**, on the other hand, is a commercial product specialized in handling PowerPoint slides, and, as such, it provides very efficient and fast ways to split, merge etc. Ti use Aspose and avoid Watermarks on slides, you must register a valid license key file using the `Conversion.SetAsposeSlidesLicense` operation

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
  "EmbeddedFonts": [],
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

    git clone https://github.com/nuxeo-powerpoint-utils.git
    cd nuxeo-powerpoint-utils
    
    mvn clean install

# Support

**These features are not part of the Nuxeo Production platform, they are not supportes**

These solutions are provided for inspiration and we encourage customers to use them as code samples and learning resources.

This is a moving project (no API maintenance, no deprecation process, etc.) If any of these solutions are found to be useful for the Nuxeo Platform in general, they will be integrated directly into platform, not maintained here.


# Licensing

[Apache License, Version 2.0](http://www.apache.org/licenses/LICENSE-2.0)


# About Nuxeo

Nuxeo dramatically improves how content-based applications are built, managed and deployed, making customers more agile, innovative and successful. Nuxeo provides a next generation, enterprise ready platform for building traditional and cutting-edge content oriented applications. Combining a powerful application development environment with SaaS-based tools and a modular architecture, the Nuxeo Platform and Products provide clear business value to some of the most recognizable brands including Verizon, Electronic Arts, Sharp, FICO, the U.S. Navy, and Boeing. Nuxeo is headquartered in New York and Paris.

More information is available at [www.nuxeo.com](http://www.nuxeo.com).  