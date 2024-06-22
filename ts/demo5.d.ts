//----&&----start----&&----
export interface demo5Data {
        id: number;
    name: string;
    zhanli: number;
    goods: number[];
}

export interface demo5DataMap3 {
    [key: string]: demo5Data;
}

export interface demo5DataMap2 {
    [key: string]: demo5DataMap3;
}

export interface demo5DataMap {
    [key: string]: demo5DataMap2;
}
//----&&----end----&&----