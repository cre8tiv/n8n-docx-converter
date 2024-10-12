import { IExecuteFunctions } from 'n8n-workflow';
import { IDataObject, INodeExecutionData, INodeType, INodeTypeDescription, NodeOperationError } from 'n8n-workflow';
import * as fs from 'fs';
import * as path from 'path';
import * as mammoth from 'mammoth';

export class DocxToText implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'DOCX to Text',
        name: 'docxToText',
        group: ['transform'],
        version: 1,
        description: 'Converts DOCX file to plain text',
        defaults: {
            name: 'DOCX to Text',
        },
        inputs: ['main'],
        outputs: ['main'],
        properties: [
            {
                displayName: 'DOCX File',
                name: 'docxFile',
                type: 'string',
                default: '',
                placeholder: 'Path to the DOCX file',
                description: 'The path to the DOCX file that you want to convert to text',
                required: true,
            },
        ],
    };

    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
        const items = this.getInputData();
        const returnData: IDataObject[] = [];

        for (let i = 0; i < items.length; i++) {
            const docxFilePath = this.getNodeParameter('docxFile', i) as string;

            if (!fs.existsSync(docxFilePath)) {
                throw new NodeOperationError(this.getNode(), `File not found: ${docxFilePath}`);
            }

            try {
                // Convert DOCX to text
                const result = await mammoth.extractRawText({ path: docxFilePath });
                const text = result.value; // The raw text

                // Save the text to an output file
                const outputFilePath = path.join(path.dirname(docxFilePath), 'output.txt');
                fs.writeFileSync(outputFilePath, text);

                returnData.push({
                    json: {
                        docxFilePath,
                        outputFilePath,
                    },
                });
            } catch (error) {
                throw new NodeOperationError(this.getNode(), `Error processing DOCX file: ${error.message}`);
            }
        }

        return [this.helpers.returnJsonArray(returnData)];
    }
}
