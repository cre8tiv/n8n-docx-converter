import {
	ICredentialType,
	INodeProperties,
} from 'n8n-workflow';

export class DocxToTextCredentialsApi implements ICredentialType {
	name = 'docxToTextCredentialsApi';
	displayName = 'DocxToText Credentials API';
	properties: INodeProperties[] = [];
}
