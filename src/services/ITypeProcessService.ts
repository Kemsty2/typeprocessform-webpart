export interface ITypeProcessService {
  createTypeProcess: (webUrl: string, listUrl, data: any) => Promise<any>;
}
