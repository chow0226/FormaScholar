import { beforeEach, describe, expect, it } from 'vitest';
import strategyJson from '../src/strategies/aip-2023.json';
import { StrategyDefinition } from '../src/core/models';
import { BUILT_IN_STRATEGY_ID, StrategyStore, validateStrategySchema } from '../src/core/StrategyStore';
import { TemplateAnalysisResult } from '../src/template/TemplateAnalyzer';

class MemoryStorage implements Storage {
  private values = new Map<string, string>();
  get length(): number { return this.values.size; }
  clear(): void { this.values.clear(); }
  getItem(key: string): string | null { return this.values.get(key) ?? null; }
  key(index: number): string | null { return Array.from(this.values.keys())[index] ?? null; }
  removeItem(key: string): void { this.values.delete(key); }
  setItem(key: string, value: string): void { this.values.set(key, value); }
}

const baseStrategy = strategyJson as StrategyDefinition;

function analysis(): TemplateAnalysisResult {
  return {
    fileName: 'AIP-2027.docx',
    components: [],
    candidates: [],
    proposals: [{
      ruleId: 'TITLE_SIZE_001',
      ruleName: 'Title size',
      component: 'title',
      property: 'font.size',
      currentValue: 18,
      detectedValue: 16,
      status: 'changed',
      confidence: 'high',
      note: 'Directly measured.'
    }],
    counts: { unchanged: 0, changed: 1, new: 0, review: 0, notDeterminable: 0 },
    warnings: []
  };
}

describe('StrategyStore', () => {
  beforeEach(() => {
    Object.defineProperty(globalThis, 'localStorage', {
      configurable: true,
      value: new MemoryStorage()
    });
  });

  it('creates imported strategies from the immutable built-in rules', () => {
    const builtIn = StrategyStore.getBuiltInStrategy();
    const imported = StrategyStore.createImportedStrategy(analysis(), new Set(['TITLE_SIZE_001']));

    expect(imported.origin).toBe('imported');
    expect(imported.baseStrategyId).toBe(BUILT_IN_STRATEGY_ID);
    expect(imported.rules.find(rule => rule.id === 'TITLE_SIZE_001')?.expected).toBe(16);
    expect(builtIn.rules.find(rule => rule.id === 'TITLE_SIZE_001')?.expected).toBe(18);
    expect(baseStrategy.rules.find(rule => rule.id === 'TITLE_SIZE_001')?.expected).toBe(18);
  });

  it('persists imported strategy activation and restores default without deleting it', () => {
    const imported = StrategyStore.createImportedStrategy(analysis(), new Set(['TITLE_SIZE_001']));
    StrategyStore.saveImportedStrategy(imported);
    StrategyStore.setActiveStrategy(imported.id);
    expect(StrategyStore.getActiveStrategy().id).toBe(imported.id);

    StrategyStore.restoreBuiltInDefault();
    expect(StrategyStore.getActiveStrategy().id).toBe(BUILT_IN_STRATEGY_ID);
    expect(StrategyStore.readState().importedStrategy?.id).toBe(imported.id);

    StrategyStore.removeImportedStrategy(imported.id);
    expect(StrategyStore.readState().importedStrategy).toBeNull();
    expect(StrategyStore.getActiveStrategy().id).toBe(BUILT_IN_STRATEGY_ID);
  });

  it('never allows the built-in strategy to be deleted', () => {
    expect(() => StrategyStore.removeImportedStrategy(BUILT_IN_STRATEGY_ID))
      .toThrow('built-in AIP Default strategy cannot be deleted');
  });

  it('deletes only the requested imported strategy', () => {
    const imported = StrategyStore.createImportedStrategy(analysis(), new Set(['TITLE_SIZE_001']));
    StrategyStore.saveImportedStrategy(imported);

    expect(() => StrategyStore.removeImportedStrategy('aip-imported-missing'))
      .toThrow('uploaded strategy is no longer available');
    expect(StrategyStore.readState().importedStrategy?.id).toBe(imported.id);
  });

  it('fails closed on corrupt imported strategy storage', () => {
    localStorage.setItem('formascholar.importedStrategy', '{not-json');
    localStorage.setItem('formascholar.activeStrategyId', 'bad-imported-id');
    const state = StrategyStore.readState();

    expect(state.activeStrategyId).toBe(BUILT_IN_STRATEGY_ID);
    expect(state.importedStrategy).toBeNull();
    expect(state.warning).toContain('could not be loaded');
  });

  it('rejects imported strategies with a reserved id or non-built-in base', () => {
    const imported = StrategyStore.createImportedStrategy(analysis(), new Set(['TITLE_SIZE_001']));

    localStorage.setItem('formascholar.importedStrategy', JSON.stringify({
      ...imported,
      id: BUILT_IN_STRATEGY_ID
    }));
    localStorage.setItem('formascholar.activeStrategyId', BUILT_IN_STRATEGY_ID);
    const reservedState = StrategyStore.readState();
    expect(reservedState.importedStrategy).toBeNull();
    expect(reservedState.activeStrategyId).toBe(BUILT_IN_STRATEGY_ID);

    localStorage.setItem('formascholar.importedStrategy', JSON.stringify({
      ...imported,
      baseStrategyId: 'another-strategy'
    }));
    const wrongBaseState = StrategyStore.readState();
    expect(wrongBaseState.importedStrategy).toBeNull();
    expect(wrongBaseState.activeStrategyId).toBe(BUILT_IN_STRATEGY_ID);

    expect(() => StrategyStore.saveImportedStrategy({
      ...imported,
      id: BUILT_IN_STRATEGY_ID
    })).toThrow('invalid imported strategy');
  });

  it('validates rule ids and rule schema before persistence', () => {
    expect(validateStrategySchema(StrategyStore.getBuiltInStrategy())).toBe(true);
    expect(validateStrategySchema({ ...baseStrategy, rules: [] })).toBe(false);
    expect(validateStrategySchema({ ...baseStrategy, rules: [baseStrategy.rules[0], baseStrategy.rules[0]] })).toBe(false);
  });
});
