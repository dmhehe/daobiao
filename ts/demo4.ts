//****&&*****start****&&****
export interface demo4Data {
        id: number;
    name: string;
    zhanli: number;
    goods: number[];
}

export type demo4DataArray = demo4Data[];

export interface demo4DataMap {
    [key: string]: demo4DataArray;
}
//****&&*****end****&&****