//----&&----start----&&----
export interface demo6Data {
        id: number;
    name: string;
    zhanli: number;
    goods: number[];
}

export type demo6DataArray = demo6Data[];

export interface demo6DataMap2 {
    [key: string]: demo6DataArray;
}

export interface demo6DataMap {
    [key: string]: demo6DataMap2;
}
//----&&----end----&&----