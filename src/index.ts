import {
  JupyterFrontEnd,
  JupyterFrontEndPlugin
} from '@jupyterlab/application';

import { DocReaderFactory } from './widget';

/**
 * File types for document reader
 */
const FILE_TYPES = [
  {
    name: 'docx',
    displayName: 'Word Document (DOCX)',
    extensions: ['.docx'],
    mimeTypes: [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    ],
    contentType: 'file',
    fileFormat: 'base64'
  },
  {
    name: 'doc',
    displayName: 'Word Document (DOC)',
    extensions: ['.doc'],
    mimeTypes: ['application/msword'],
    contentType: 'file',
    fileFormat: 'base64'
  },
  {
    name: 'rtf',
    displayName: 'Rich Text Format',
    extensions: ['.rtf'],
    mimeTypes: ['application/rtf', 'text/rtf'],
    contentType: 'file',
    fileFormat: 'base64'
  },
  {
    name: 'pptx',
    displayName: 'PowerPoint Presentation (PPTX)',
    extensions: ['.pptx'],
    mimeTypes: [
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    ],
    contentType: 'file',
    fileFormat: 'base64'
  },
  {
    name: 'ppt',
    displayName: 'PowerPoint Presentation (PPT)',
    extensions: ['.ppt'],
    mimeTypes: ['application/vnd.ms-powerpoint'],
    contentType: 'file',
    fileFormat: 'base64'
  },
  {
    name: 'odt',
    displayName: 'OpenDocument Text (ODT)',
    extensions: ['.odt'],
    mimeTypes: ['application/vnd.oasis.opendocument.text'],
    contentType: 'file',
    fileFormat: 'base64'
  },
  {
    name: 'odp',
    displayName: 'OpenDocument Presentation (ODP)',
    extensions: ['.odp'],
    mimeTypes: ['application/vnd.oasis.opendocument.presentation'],
    contentType: 'file',
    fileFormat: 'base64'
  },
  {
    name: 'xlsx',
    displayName: 'Excel Workbook (XLSX)',
    extensions: ['.xlsx'],
    mimeTypes: [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ],
    contentType: 'file',
    fileFormat: 'base64'
  }
];

/**
 * Initialization data for the jupyterlab_doc_reader_extension extension.
 */
const plugin: JupyterFrontEndPlugin<void> = {
  id: 'jupyterlab_doc_reader_extension:plugin',
  description:
    'Renders DOCX, PPTX, RTF, ODT, ODP and XLSX documents in the browser',
  autoStart: true,
  activate: (app: JupyterFrontEnd) => {
    console.log(
      'JupyterLab extension jupyterlab_doc_reader_extension is activated!'
    );

    const { docRegistry } = app;

    // Register file types - mark them as binary to prevent text loading
    FILE_TYPES.forEach(fileType => {
      try {
        docRegistry.addFileType({
          name: fileType.name,
          displayName: fileType.displayName,
          extensions: fileType.extensions,
          mimeTypes: fileType.mimeTypes
        });
        console.log(`Registered file type: ${fileType.name}`);
      } catch (e) {
        console.warn(`File type ${fileType.name} already registered`, e);
      }
    });

    // Create widget factory - use base64 model to handle binary files
    const factory = new DocReaderFactory({
      name: 'Document Reader',
      modelName: 'base64',
      fileTypes: FILE_TYPES.map(ft => ft.name),
      defaultFor: FILE_TYPES.map(ft => ft.name),
      readOnly: true
    });

    // Register the factory
    docRegistry.addWidgetFactory(factory);

    console.log(
      'Document reader widget factory registered for:',
      FILE_TYPES.map(ft => ft.name).join(', ')
    );
  }
};

export default plugin;
