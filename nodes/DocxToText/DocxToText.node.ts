import { IExecuteFunctions } from 'n8n-workflow';
import { IDataObject, INodeExecutionData, INodeType, INodeTypeDescription, NodeOperationError } from 'n8n-workflow';
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
                displayName: 'Input Binary Field',
                name: 'inputBinaryField',
                type: 'string',
                default: 'data',
                placeholder: 'Input binary field containing the DOCX file',
                description: 'The name of the input binary field containing the DOCX file',
                required: true,
            },
            {
                displayName: 'Destination Output Field',
                name: 'destinationOutputField',
                type: 'string',
                default: 'text',
                placeholder: 'Destination output field for the converted text',
                description: 'The name of the destination output field for the converted text',
                required: true,
            },
        ],
    };

    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
        const items = this.getInputData();
        const returnData: IDataObject[] = [];

        for (let i = 0; i < items.length; i++) {
            const inputBinaryField = this.getNodeParameter('inputBinaryField', i) as string;
            const destinationOutputField = this.getNodeParameter('destinationOutputField', i) as string;

            const binaryData = await this.helpers.getBinaryDataBuffer(i, inputBinaryField);
            if (!binaryData) {
                throw new NodeOperationError(this.getNode(), `No binary data found for field "${inputBinaryField}"`);
            }

            const result = await mammoth.extractRawText({ buffer: binaryData });
            const text = result.value;

            returnData.push({
                json: {
                    [destinationOutputField]: text,
                },
            });
        }

        return [this.helpers.returnJsonArray(returnData)];
    }
}
