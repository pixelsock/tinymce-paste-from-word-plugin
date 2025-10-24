/**
 * Copyright (c) Tiny Technologies, Inc. and Pangaea Information Technologies, Ltd.
 * All rights reserved.
 * Licensed under the LGPL or a commercial license.
 * For LGPL see License.txt in the project root for license information.
 * For commercial licenses see https://www.tiny.cloud/
 */
import tinymce, { Editor } from "tinymce";
import { isWordContent, preProcess, PreProcessEvent, PostProcessEvent } from "./WordFilter";

export default (): void => {
  tinymce.PluginManager.add("paste_from_word", (editor: Editor) => {
    editor.on("PastePreProcess", (args: PreProcessEvent) => {
      // Ensure mode and source are set properly
      if (!args.mode) {
        args.mode = "html";
      }
      if (!args.source) {
        args.source = args.internal ? "internal" : "external";
      }

      // Fire custom pre-process event for Word content
      if (isWordContent(args.content)) {
        const preProcessData = {
          content: args.content,
          mode: args.mode,
          source: args.source
        };
        
        editor.dispatch("PasteFromWordPreProcess", preProcessData);
        
        // Process the content
        args.content = preProcess(editor, args.content);
      }
    });

    editor.on("PastePostProcess", (args: PostProcessEvent) => {
      // Ensure mode and source are set properly
      if (!args.mode) {
        args.mode = "html";
      }
      if (!args.source) {
        args.source = args.internal ? "internal" : "external";
      }

      // Fire custom post-process event for Word content
      const postProcessData = {
        node: args.node,
        mode: args.mode,
        source: args.source
      };
      
      editor.dispatch("PasteFromWordPostProcess", postProcessData);
    });

    return {
      getMetadata: () => ({
        name: "Paste from Word",
        url: "https://github.com/pixelsock/tinymce-paste-from-word-plugin",
      }),
    };
  });
};
