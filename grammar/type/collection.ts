/**
 * Represents unions.
 * (A1, A1:C5, ...)
 */
class Collection {
    private _data: any[];
    private _refs: any[];

    constructor(data?: any[], refs?: any[]) {
        if (data == null && refs == null) {
            this._data = [];
            this._refs = [];
        } else {
            if (data && refs && data.length !== refs.length) {
                throw Error('Collection: data length should match references length.');
            }
            this._data = data || [];
            this._refs = refs || [];
        }
    }

    get data(): any[] {
        return this._data;
    }

    get refs(): any[] {
        return this._refs;
    }

    get length(): number {
        return this._data.length;
    }

    /**
     * Add data and references to this collection.
     * @param obj - data
     * @param ref - reference
     */
    add(obj: any, ref: any): void {
        this._data.push(obj);
        this._refs.push(ref);
    }
}

export default Collection;