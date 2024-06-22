//----&&----start----&&----
export interface demo3Data {
        id: number;
    name: string;
    zhanli: number;
    goods: number[];
}

export interface demo3DataMap2 {
    [key: string]: demo3Data;
}

export interface demo3DataMap {
    [key: string]: demo3DataMap2;
}
//----&&----end----&&----