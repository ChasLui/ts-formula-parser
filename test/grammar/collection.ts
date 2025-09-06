import { expect, describe, it } from 'vitest';
import Collection from '../../grammar/type/collection';
import type { CellRef } from '../../index.d';

describe('Collection', () => {
    it('should throw error', function () {
        const refs: CellRef[] = [{row: 1, col: 1, sheet: 'Sheet1'}];
        expect((() => new Collection([], refs)))
            .to.throw('Collection: data length should match references length.')
    });

    it('should not throw error', function () {
        const refs: CellRef[] = [{row: 1, col: 1, sheet: 'Sheet1'}];
        expect((() => new Collection([1], refs)))
            .to.not.throw()
    });
});