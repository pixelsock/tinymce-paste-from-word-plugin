# TinyMCE Paste from Word Plugin

> **Note:** This is a continuation of the [original archived project](https://github.com/pangaeatech/tinymce-paste-from-word-plugin). I've updated and maintained this plugin to support the latest versions of TinyMCE (6.x, 7.x, and 8.x) with enhanced features including image support and custom event handling.

A free, open-source plugin that provides paste-from-Word functionality for TinyMCE 6+. This plugin cleans up messy HTML from Microsoft Word documents, preserves formatting, and now supports embedded images.

## Why Use This Plugin?

- 🆓 **Free Alternative** to TinyMCE's premium PowerPaste plugin
- 📝 **Clean Paste** - Automatically cleans up Word's messy HTML
- 🖼️ **Image Support** - Paste images from Word documents with Base64 encoding
- 🎨 **Custom Events** - Hook into paste events for custom formatting
- 🚀 **Modern** - Works with TinyMCE 6.x, 7.x, and 8.x
- ✅ **Tested** - Comprehensive test coverage and actively maintained

## Features

* **TinyMCE 8.x Compatibility:** Full support for TinyMCE 6.x, 7.x, and 8.x
* **Image Support:** Paste images from Word documents using `paste_data_images`
* **Custom Events:** `PasteFromWordPreProcess` and `PasteFromWordPostProcess` events for custom formatting
* **Base64 Image Handling:** Automatic Base64 encoding with `images_upload_handler`
* **Smart List Conversion:** Converts Word's fake lists to proper HTML lists
* **Style Preservation:** Keeps important formatting while removing Word junk

## Comparison with PowerPaste

| Feature                           | This Plugin | PowerPaste |
| :-------------------------------- | :---------: | :--------: |
| **Cost**                          |    Free     |    Paid    |
| Automatically cleans up content   |     ✔      |     ✔     |
| Supports embedded images          |     ✔      |     ✔     |
| Paste from Microsoft Word         |     ✔      |     ✔     |
| Paste from Microsoft Word online  |     ✔      |     ✔     |
| Paste from Microsoft Excel        |      -      |     ✔     |
| Paste from Microsoft Excel online |      -      |     -      |
| Paste from Google Docs            |     ✔      |     ✔     |
| Paste from Google Sheets          |      -      |     -      |
| Custom paste events               |     ✔      |     -      |
| Open source                       |     ✔      |     -      |

## Quick Start

### Installation

Choose one of the following installation methods:

#### Option 1: CDN (Easiest)

Load the plugin directly from unpkg CDN:

```js
tinymce.PluginManager.load(
  "paste_from_word",
  "https://unpkg.com/@pixelsock/tinymce-paste-from-word-plugin@latest/index.js",
);
tinymce.init({
  selector: "textarea", // change this value according to your HTML
  plugins: "paste_from_word image",
  paste_webkit_styles: "all",
  paste_remove_styles_if_webkit: false,
  paste_data_images: true, // Enable image pasting as Base64
  images_upload_handler: function(blobInfo, success, failure) {
    // Handle image upload or convert to Base64
    const reader = new FileReader();
    reader.onload = function() {
      success(reader.result);
    };
    reader.readAsDataURL(blobInfo.blob());
  }
});
```

#### Option 2: npm/yarn

Install via npm or yarn:

```bash
npm install @pixelsock/tinymce-paste-from-word-plugin
# or
yarn add @pixelsock/tinymce-paste-from-word-plugin
```

Then load it in your project (see React example below).

#### Option 3: Self-Hosted

1. Create a new folder `paste_from_word` inside of the existing TinyMCE `plugins` folder
2. Download [`index.js`](https://unpkg.com/@pixelsock/tinymce-paste-from-word-plugin@latest/index.js) from unpkg
3. Save it to the new folder as `plugin.min.js`
4. Configure your TinyMCE instance to use the plugin:

```js
tinymce.init({
  selector: "textarea", // change this value according to your HTML
  plugins: "paste_from_word image",
  paste_webkit_styles: "all",
  paste_remove_styles_if_webkit: false,
  paste_data_images: true, // Enable image pasting as Base64
  images_upload_handler: function(blobInfo, success, failure) {
    // Handle image upload or convert to Base64
    const reader = new FileReader();
    reader.onload = function() {
      success(reader.result);
    };
    reader.readAsDataURL(blobInfo.blob());
  }
});
```

### Usage with React

For React projects (or other Node.js frameworks):

1. Install the required packages:

```bash
npm install @tinymce/tinymce-react @pixelsock/tinymce-paste-from-word-lib
```

2. Use in your React component:

```jsx
import React from "react";
import { Editor } from "@tinymce/tinymce-react";
import PasteFromWord from "@pixelsock/tinymce-paste-from-word-lib";

const config = {
  height: 500,
  plugins: "image",
  paste_preprocess: PasteFromWord,
  paste_webkit_styles: "all",
  paste_remove_styles_if_webkit: false,
  paste_data_images: true,
  images_upload_handler: (blobInfo, success, failure) => {
    const reader = new FileReader();
    reader.onload = () => success(reader.result);
    reader.onerror = () => failure('Image upload failed');
    reader.readAsDataURL(blobInfo.blob());
  },
  setup: (editor) => {
    editor.on('PasteFromWordPreProcess', (e) => {
      console.log('Pre-process:', e);
    });
    editor.on('PasteFromWordPostProcess', (e) => {
      console.log('Post-process:', e);
    });
  }
};

export default function App() {
  return (
    <Editor
      initialValue="<p>This is the initial content of the editor.</p>"
      init={config}
    />
  );
}
```

## Events

This plugin fires custom events that you can listen to for advanced customization.

### `PasteFromWordPreProcess`

Fired before the content is parsed from Word documents.

**Event Data:**
- `content` (string): The HTML content being pasted
- `mode` (string): The paste mode (usually "html")
- `source` (string): The paste source ("internal" or "external")

**Example:**
```js
tinymce.init({
  selector: "textarea",
  plugins: "paste_from_word",
  setup: function(editor) {
    editor.on('PasteFromWordPreProcess', function(e) {
      console.log('Content before processing:', e.content);
      // Modify content before processing
      e.content = e.content.replace(/something/, 'something else');
    });
  }
});
```

### `PasteFromWordPostProcess`

Fired after the content has been parsed from Word documents, but before it's added to the editor.

**Event Data:**
- `node` (Element): The DOM node containing the processed content
- `mode` (string): The paste mode (usually "html")
- `source` (string): The paste source ("internal" or "external")

**Example:**
```js
tinymce.init({
  selector: "textarea",
  plugins: "paste_from_word",
  setup: function(editor) {
    editor.on('PasteFromWordPostProcess', function(e) {
      console.log('Processed node:', e.node);
      // Modify the DOM after processing
      e.node.setAttribute('data-pasted-from-word', 'true');
    });
  }
});
```

## Settings

These settings affect the execution of the `paste_from_word` plugin.

### `pastefromword_valid_elements`

This option enables you to configure the elements specific to MS Office. Word produces a lot of junk HTML, so when users paste things from Word we do extra restrictive filtering on it to remove as much of this as possible. This option enables you to specify which elements and attributes you want to include when Word contents are intercepted.

Type: String

Default Value: `"-strong/b,-em/i,-u,-span,-p,-ol,-ul,-li,-h1,-h2,-h3,-h4,-h5,-h6,-p/div,-a[href|name],sub,sup,strike,br,del,table[width],tr,td[colspan|rowspan|width],th[colspan|rowspan|width],thead,tfoot,tbody"`

### `pastefromword_convert_fake_lists`

This option lets you disable the logic that converts list like paragraph structures into real semantic HTML lists.

Type: Boolean

Default Value: `true`

### `paste_data_images`

When enabled, images pasted from Word and other sources will be imported as Base64 encoded images. This allows images to be embedded directly in the content without requiring server-side storage. You can then use `images_upload_handler` to automatically upload these Base64 images to your server if desired.

Type: Boolean

Default Value: `false`

### `images_upload_handler`

A custom function that handles image uploads. When images are pasted (with `paste_data_images: true`), this handler is called. You can convert images to Base64 or upload them to a server and return a URL.

Type: Function

Default Value: `undefined`

**Example:**
```js
images_upload_handler: function(blobInfo, success, failure) {
  // Option 1: Convert to Base64
  const reader = new FileReader();
  reader.onload = function() {
    success(reader.result);
  };
  reader.readAsDataURL(blobInfo.blob());
  
  // Option 2: Upload to server (example)
  // const formData = new FormData();
  // formData.append('file', blobInfo.blob(), blobInfo.filename());
  // fetch('/upload', { method: 'POST', body: formData })
  //   .then(response => response.json())
  //   .then(data => success(data.location))
  //   .catch(error => failure('Upload failed'));
}
```

### `paste_webkit_styles`

This plugin is a preprocessor which converts paste content from MS Word into WebKit-style paste content which TinyMCE's built-in paste function can handle. Therefore, it is impacted by the webkit-specific settings of the paste module. In order to prevent the paste module from stripping out all style information, you need to set this to `"all"` or to a specific list of styles you wish to retain.

Type: String

Default value: `"none"`

### `paste_remove_styles_if_webkit`

This plugin is a preprocessor which converts paste content from MS Word into WebKit-style paste content which TinyMCE's built-in paste function can handle. Therefore, it is impacted by the webkit-specific settings of the paste module. In order to prevent the paste module from stripping out all style information, you need to set this to `false`.

Type: Boolean

Default Value: `true`

## Project History & Maintenance

This project is a continuation and enhancement of the [original TinyMCE Paste from Word plugin](https://github.com/pangaeatech/tinymce-paste-from-word-plugin) by Pangaea Information Technologies, Ltd., which was archived when TinyMCE 6.x support ended in October 2024.

### What's New in This Fork

I've revived and updated this plugin with:

- ✅ **TinyMCE 8.x Support** - Full compatibility with TinyMCE 6.x, 7.x, and 8.x
- ✅ **Image Support** - Paste images from Word with Base64 encoding
- ✅ **Custom Events** - New `PasteFromWordPreProcess` and `PasteFromWordPostProcess` events
- ✅ **Active Maintenance** - Regular updates and bug fixes
- ✅ **Modern Tooling** - Updated dependencies and build process

### Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

### License

This project is licensed under the LGPL-2.1 License, maintaining the same license as the original project.

### Credits

- Original plugin by [Pangaea Information Technologies, Ltd.](https://github.com/pangaeatech)
- Based on TinyMCE 5.x paste-from-word functionality
- Maintained and updated by [@pixelsock](https://github.com/pixelsock)

### Support

- 🐛 [Report Issues](https://github.com/pixelsock/tinymce-paste-from-word-plugin/issues)
- 💬 [Discussions](https://github.com/pixelsock/tinymce-paste-from-word-plugin/discussions)
- 📖 [Documentation](https://github.com/pixelsock/tinymce-paste-from-word-plugin)

---

**Made with ❤️ for the TinyMCE community**
