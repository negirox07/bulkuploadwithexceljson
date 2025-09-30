
export interface ISPHelper {
    getListData(url: string):Promise<any>;
    getListDataRecursive(url: string,data:any[]): Promise<any>;
}