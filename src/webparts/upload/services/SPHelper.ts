import { postheaders, putHeaders } from "../models/SPConstants";
import {
    SPHttpClient, SPHttpClientResponse
} from '@microsoft/sp-http';
import { ISPHelper } from "./ISPHelper";
export class SPHelpers implements ISPHelper {
    private _client: SPHttpClient;
    private _clientConfig = SPHttpClient.configurations.v1;
    constructor(client: SPHttpClient) {
        this._client = client;
    }
    public async getListData(url: string): Promise<any> {
        const response = await this._client.get(url, this._clientConfig);
        return await response.json();
    }
    public async getListDataRecursive(url: string, data: any[]): Promise<any> {
        const response = await this._client.get(url, this._clientConfig);
        const responseJson = await response.json();
        data = data.concat(responseJson.value);
        if (responseJson['@odata.nextLink']) {
            return this.getListDataRecursive(responseJson['@odata.nextLink'], data);
        }
        return data;
    }
    public async setListData(url: string, postData: string): Promise<SPHttpClientResponse> {
        const headers = { ...postheaders }
        headers.body = postData;
        const response = await this._client.post(url, this._clientConfig, headers);
        return response;
    }
    public async putListData(url: string, postData: string): Promise<SPHttpClientResponse> {
        const headers = { ...putHeaders }
        headers.body = postData;
        const response = await this._client.post(url, this._clientConfig, headers);
        return response;
    }
}