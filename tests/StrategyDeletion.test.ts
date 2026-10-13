import { describe, expect, it } from 'vitest';
import { isDeleteConfirmationCurrent } from '../src/taskpane/StrategyDeletion';

describe('strategy deletion confirmation', () => {
  it('accepts confirmation only for the same imported strategy id', () => {
    expect(isDeleteConfirmationCurrent('imported-a', 'imported-a')).toBe(true);
    expect(isDeleteConfirmationCurrent('imported-a', 'imported-b')).toBe(false);
  });

  it('rejects missing confirmation or imported strategy ids', () => {
    expect(isDeleteConfirmationCurrent(null, 'imported-a')).toBe(false);
    expect(isDeleteConfirmationCurrent('imported-a', null)).toBe(false);
    expect(isDeleteConfirmationCurrent('', 'imported-a')).toBe(false);
  });
});
