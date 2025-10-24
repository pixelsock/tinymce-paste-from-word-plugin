# TinyMCE Paste from Word Plugin

This plugin adds the open-source [Paste from Word](https://www.tiny.cloud/docs/plugins/opensource/paste/) functionality from the 5.x branch of TinyMCE as a plugin for the 6.x, 7.x, and 8.x branches. The goal of this project is not to replace the premium [PowerPaste plugin](https://www.tiny.cloud/tinymce/features/powerpaste/), but to allow users to have paste-from-word support with the latest versions of TinyMCE.

## New Features

* **TinyMCE 8.x Compatibility:** Now supports TinyMCE 6.x, 7.x, and 8.x
* **Image Support:** Paste images from Word documents using `paste_data_images`
* **Custom Events:** `PasteFromWordPreProcess` and `PasteFromWordPostProcess` events for custom formatting
* **Base64 Image Handling:** Automatic Base64 encoding with `images_upload_handler`

### Comparison with PowerPaste

| Feature                           | This Plugin | PowerPaste |
| :-------------------------------- | :---------: | :--------: |
| Automatically cleans up content   |     ✔      |     ✔     |
| Supports embedded images          |     ✔      |     ✔     |
| Paste from Microsoft Word         |     ✔      |     ✔     |
| Paste from Microsoft Word online  |     ✔      |     ✔     |
| Paste from Microsoft Excel        |      -      |     ✔     |
| Paste from Microsoft Excel online |      -      |     -      |
| Paste from Google Docs            |     ✔      |     ✔     |
| Paste from Google Sheets          |      -      |     -      |
| Custom paste events               |     ✔      |     -      |

## Usage

### Option 1: CDN Hosted

1. Tell your TinyMCE instance where to load the plugin from and how to configure it:

```js
tinymce.PluginManager.load(
  "paste_from_word",
  "https://unpkg.com/@pangaeatech/tinymce-paste-from-word-plugin@latest/index.js",
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

### Option 2: Self-Hosted

1. Create a new folder `paste_from_word` inside of the existing TinyMCE `plugins` folder.
2. Download the file `https://unpkg.com/@pangaeatech/tinymce-paste-from-word-plugin@latest/index.js` and add it to that new folder, renaming it `plugin.min.js`
3. Configure your TinyMCE instance to use the plugin:

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

### Option 3: React (etc.)

The following instructions are for a project using ReactJS and NPM, but you can
easily modify these for any other NodeJS-based project.

1. Add the TinyMCE and TinyMCE Paste from Word Plugin projects to your package management:

```bash
npx create-react-app tinymce-react-demo
cd tinymce-react-demo
npm install --save @tinymce/tinymce-react @pangaeatech/tinymce-paste-from-word-lib
```

2. Using a text editor, open ./src/App.js and replace the contents with:

```jsx
import React from "react";
import { Editor } from "@tinymce/tinymce-react";
import PasteFromWord from "@pangaeatech/tinymce-paste-from-word-lib";

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
