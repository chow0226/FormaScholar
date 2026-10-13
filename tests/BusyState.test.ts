import { describe, expect, it } from 'vitest';
import { canStartCorrection, nextBusyDepth } from '../src/taskpane/BusyState';

describe('taskpane busy-state guards', () => {
  it('keeps nested operations busy until the final release', () => {
    let depth = 0;

    depth = nextBusyDepth(depth, true);
    expect(depth).toBe(1);

    depth = nextBusyDepth(depth, true);
    expect(depth).toBe(2);

    depth = nextBusyDepth(depth, false);
    expect(depth).toBe(1);

    depth = nextBusyDepth(depth, false);
    expect(depth).toBe(0);
  });

  it('never lets busy depth become negative', () => {
    expect(nextBusyDepth(0, false)).toBe(0);
  });

  it('blocks duplicate correction starts while busy', () => {
    expect(canStartCorrection(true, true, 1)).toBe(false);
    expect(canStartCorrection(false, true, 1)).toBe(true);
    expect(canStartCorrection(false, false, 1)).toBe(false);
    expect(canStartCorrection(false, true, 0)).toBe(false);
  });
});
