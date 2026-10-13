import aipStrategy from '../strategies/aip-2023.json';
import { StrategyDefinition } from './models';
import { TemplateAnalysisResult, TemplateAnalyzer } from '../template/TemplateAnalyzer';

export const BUILT_IN_STRATEGY_ID = 'aip-default';
const IMPORTED_STORAGE_KEY = 'formascholar.importedStrategy';
const ACTIVE_STORAGE_KEY = 'formascholar.activeStrategyId';

const VALID_OPERATORS = new Set(['equals', 'not_equals', 'required', 'forbidden', 'matches_regex', 'maximum', 'minimum', 'sequence']);
const VALID_SEVERITIES = new Set(['critical', 'major', 'minor', 'info']);
const VALID_TARGETS = new Set(['document', 'title', 'authorName', 'authorAffiliation', 'authorEmail', 'abstract', 'body', 'heading1', 'heading2', 'heading3', 'figureCaption', 'tableCaption', 'equation', 'referencesHeading', 'reference']);

export interface StrategyStoreState {
  activeStrategyId: string;
  importedStrategy: StrategyDefinition | null;
  warning?: string;
}

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function deepFreeze<T>(value: T): T {
  if (!value || typeof value !== 'object' || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child);
  return value;
}

function storage(): Storage | null {
  try {
    return typeof localStorage === 'undefined' ? null : localStorage;
  } catch {
    return null;
  }
}

export function validateStrategySchema(value: unknown): value is StrategyDefinition {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as Partial<StrategyDefinition>;
  if (typeof candidate.id !== 'string' || !candidate.id.trim()) return false;
  if (typeof candidate.name !== 'string' || typeof candidate.version !== 'string' || typeof candidate.source !== 'string') return false;
  if (!Array.isArray(candidate.rules) || candidate.rules.length === 0) return false;
  const ids = new Set<string>();
  for (const rule of candidate.rules) {
    if (!rule || typeof rule !== 'object') return false;
    if (typeof rule.id !== 'string' || ids.has(rule.id)) return false;
    if (typeof rule.name !== 'string' || typeof rule.category !== 'string') return false;
    if (typeof rule.target !== 'string' || !VALID_TARGETS.has(rule.target)) return false;
    if (typeof rule.property !== 'string' || typeof rule.operator !== 'string' || !VALID_OPERATORS.has(rule.operator)) return false;
    if (typeof rule.severity !== 'string' || !VALID_SEVERITIES.has(rule.severity)) return false;
    if (typeof rule.autoFixable !== 'boolean' || typeof rule.implemented !== 'boolean') return false;
    if (typeof rule.evidence !== 'string') return false;
    if (rule.assessmentRole !== undefined && rule.assessmentRole !== 'compliance' && rule.assessmentRole !== 'consistency') return false;
    if (rule.reviewMode !== undefined && rule.reviewMode !== 'manual' && rule.reviewMode !== 'submission' && rule.reviewMode !== 'technical') return false;
    if (rule.expected !== undefined && !['string', 'number', 'boolean'].includes(typeof rule.expected)) return false;
    ids.add(rule.id);
  }
  return true;
}

function isValidImportedStrategy(value: unknown): value is StrategyDefinition {
  if (!validateStrategySchema(value)) return false;
  const strategy = value as StrategyDefinition;
  return strategy.origin === 'imported'
    && strategy.id !== BUILT_IN_STRATEGY_ID
    && strategy.baseStrategyId === BUILT_IN_STRATEGY_ID
    && typeof strategy.sourceTemplateName === 'string'
    && strategy.sourceTemplateName.trim().length > 0
    && typeof strategy.importedAt === 'string'
    && strategy.importedAt.trim().length > 0;
}

export class StrategyStore {
  static getBuiltInStrategy(): StrategyDefinition {
    const builtIn = clone(aipStrategy as StrategyDefinition);
    builtIn.id = BUILT_IN_STRATEGY_ID;
    builtIn.origin = 'built-in';
    return deepFreeze(builtIn);
  }

  static readState(): StrategyStoreState {
    const builtIn = this.getBuiltInStrategy();
    const store = storage();
    if (!store) return { activeStrategyId: builtIn.id, importedStrategy: null };

    let importedStrategy: StrategyDefinition | null = null;
    let warning: string | undefined;
    try {
      const raw = store.getItem(IMPORTED_STORAGE_KEY);
      if (raw) {
        const parsed: unknown = JSON.parse(raw);
        if (isValidImportedStrategy(parsed)) {
          importedStrategy = clone(parsed);
        } else {
          store.removeItem(IMPORTED_STORAGE_KEY);
          warning = 'The saved imported strategy was invalid and has been discarded. AIP Default remains available.';
        }
      }
    } catch {
      store.removeItem(IMPORTED_STORAGE_KEY);
      warning = 'The saved imported strategy could not be loaded. AIP Default remains available.';
    }

    const requestedActive = store.getItem(ACTIVE_STORAGE_KEY) || builtIn.id;
    const activeStrategyId = requestedActive === importedStrategy?.id ? requestedActive : builtIn.id;
    if (activeStrategyId !== requestedActive) {
      try { store.setItem(ACTIVE_STORAGE_KEY, builtIn.id); } catch { /* storage is optional */ }
    }
    return { activeStrategyId, importedStrategy, warning };
  }

  static getActiveStrategy(): StrategyDefinition {
    const state = this.readState();
    return state.activeStrategyId === state.importedStrategy?.id && state.importedStrategy
      ? clone(state.importedStrategy)
      : this.getBuiltInStrategy();
  }

  static getAvailableStrategies(): StrategyDefinition[] {
    const state = this.readState();
    return [this.getBuiltInStrategy(), ...(state.importedStrategy ? [clone(state.importedStrategy)] : [])];
  }

  static saveImportedStrategy(strategy: StrategyDefinition): void {
    if (!isValidImportedStrategy(strategy)) throw new Error('Cannot save an invalid imported strategy.');
    const store = storage();
    if (!store) throw new Error('Browser storage is unavailable; the imported strategy could not be saved.');
    store.setItem(IMPORTED_STORAGE_KEY, JSON.stringify(clone(strategy)));
  }

  static setActiveStrategy(strategyId: string): void {
    const available = this.getAvailableStrategies();
    if (!available.some(strategy => strategy.id === strategyId)) throw new Error('That strategy is no longer available.');
    const store = storage();
    if (!store) throw new Error('Browser storage is unavailable; the active strategy could not be changed.');
    store.setItem(ACTIVE_STORAGE_KEY, strategyId);
  }

  static restoreBuiltInDefault(): void {
    const store = storage();
    if (!store) return;
    store.setItem(ACTIVE_STORAGE_KEY, BUILT_IN_STRATEGY_ID);
  }

  static removeImportedStrategy(strategyId: string): void {
    if (strategyId === BUILT_IN_STRATEGY_ID) {
      throw new Error('The built-in AIP Default strategy cannot be deleted.');
    }
    const store = storage();
    if (!store) throw new Error('Browser storage is unavailable; the imported strategy could not be deleted.');
    const state = this.readState();
    if (!state.importedStrategy || state.importedStrategy.id !== strategyId) {
      throw new Error('The uploaded strategy is no longer available.');
    }
    store.removeItem(IMPORTED_STORAGE_KEY);
    if (store.getItem(ACTIVE_STORAGE_KEY) === strategyId) {
      store.setItem(ACTIVE_STORAGE_KEY, BUILT_IN_STRATEGY_ID);
    }
  }

  static createImportedStrategy(
    analysis: TemplateAnalysisResult,
    acceptedRuleIds: Set<string>
  ): StrategyDefinition {
    const base = this.getBuiltInStrategy();
    const imported = TemplateAnalyzer.applyAcceptedChanges(base, analysis, acceptedRuleIds);
    const id = `aip-imported-${Date.now()}`;
    return {
      ...imported,
      id,
      name: `AIP – Imported from ${analysis.fileName}`,
      version: `imported-${new Date().toISOString().slice(0, 10)}`,
      origin: 'imported',
      sourceTemplateName: analysis.fileName,
      importedAt: new Date().toISOString(),
      baseStrategyId: BUILT_IN_STRATEGY_ID,
      approvedChangeCount: acceptedRuleIds.size
    };
  }
}
