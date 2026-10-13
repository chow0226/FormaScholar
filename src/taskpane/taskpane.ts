import './taskpane.css';
import { CorrectionResult, DocumentSnapshot, StrategyDefinition, ValidationSummary, Finding } from '../core/models';
import { paragraphTextFingerprint, WordDocumentAdapter } from '../word/WordDocumentAdapter';
import { ValidationEngine } from '../core/ValidationEngine';
import { CorrectionService } from '../core/CorrectionService';
import { documentFingerprint } from '../core/DocumentFingerprint';
import { filterFindings } from './FindingsFilter';
import { isTemplateProposalAcceptable, safeTemplateProposalIds, StrategyChangeProposal, TemplateAnalysisResult, TemplateAnalyzer } from '../template/TemplateAnalyzer';
import { BUILT_IN_STRATEGY_ID, StrategyStore } from '../core/StrategyStore';
import { canStartCorrection, nextBusyDepth } from './BusyState';
import { isDeleteConfirmationCurrent } from './StrategyDeletion';

const initialStrategyState = StrategyStore.readState();
const builtInAipStrategy = StrategyStore.getBuiltInStrategy();
let activeStrategyId = initialStrategyState.activeStrategyId;
let activeStrategy: StrategyDefinition = StrategyStore.getActiveStrategy();
let strategyStoreWarning = initialStrategyState.warning;

// Validation workspace state. Template analysis never replaces these values.
let currentSummary: ValidationSummary | null = null;
let currentSnapshot: DocumentSnapshot | null = null;
let validatedDocumentFingerprint: string | null = null;
let validatedStrategyFingerprint: string | null = null;
let activeFilter = 'all';
const selectedRuleIds = new Set<string>();

// Template Rules workspace state. Uploading/analyzing a template never
// silently changes the active manuscript-validation result.
let templateAnalysis: TemplateAnalysisResult | null = null;
const acceptedTemplateChanges = new Set<string>();
let busyDepth = 0;
let operationBusy = false;
let wordHostAvailable = false;
type WorkspaceView = 'validation' | 'templates';
let pendingDeleteStrategyId: string | null = null;

const $ = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
const esc = (value: unknown): string => String(value ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#039;');

function strategyFingerprint(strategy: StrategyDefinition): string {
  return JSON.stringify({
    id: strategy.id,
    version: strategy.version,
    origin: strategy.origin,
    rules: strategy.rules
  });
}

function showWorkspace(view: WorkspaceView): void {
  const validationVisible = view === 'validation';
  const templatesVisible = view === 'templates';
  $('validation-view').classList.toggle('hidden', !validationVisible);
  $('validation-view').setAttribute('aria-hidden', String(!validationVisible));
  $('templates-view').classList.toggle('hidden', !templatesVisible);
  $('templates-view').setAttribute('aria-hidden', String(!templatesVisible));
  document.querySelectorAll<HTMLButtonElement>('.mode-tab').forEach(button => {
    const selected = button.dataset.view === view;
    button.classList.toggle('active', selected);
    button.setAttribute('aria-selected', String(selected));
  });
}

Office.onReady(info => {
  wordHostAvailable = info.host === Office.HostType.Word;
  refreshStrategyUi();
  if (strategyStoreWarning) {
    showError(strategyStoreWarning);
    strategyStoreWarning = undefined;
  }

  if (!wordHostAvailable) {
    showError('FormaScholar must be opened inside Microsoft Word.');
    $('status-dot').textContent = 'Word not detected';
    updateOperationControlState();
    return;
  }

  $('status-dot').textContent = 'Connected to Word';
  bindEvents();
});

function bindEvents(): void {
  document.querySelectorAll<HTMLButtonElement>('.mode-tab').forEach(button => {
    button.addEventListener('click', () => {
      if (operationBusy) return;
      const view = button.dataset.view;
      if (view === 'validation' || view === 'templates') showWorkspace(view);
    });
  });
  $('validate').addEventListener('click', validateDocument);
  $('strategy').addEventListener('change', handleStrategyChange);
  $('upload-template').addEventListener('click', () => {
    ($('template-file') as HTMLInputElement).click();
  });
  $('template-file').addEventListener('change', () => {
    const file = ($('template-file') as HTMLInputElement).files?.[0];
    if (file) void analyzeUploadedTemplate(file);
  });
  $('select-all-template-changes').addEventListener('click', () => {
    if (operationBusy || !templateAnalysis) return;
    safeTemplateProposalIds(templateAnalysis).forEach(ruleId => acceptedTemplateChanges.add(ruleId));
    renderTemplateAnalysis();
  });
  $('clear-template-changes').addEventListener('click', () => {
    if (operationBusy) return;
    acceptedTemplateChanges.clear();
    renderTemplateAnalysis();
  });
  $('download-template-strategy').addEventListener('click', downloadApprovedTemplateStrategy);
  $('create-imported-strategy').addEventListener('click', createImportedStrategy);
  $('use-imported-strategy').addEventListener('click', () => switchStrategyTo(StrategyStore.readState().importedStrategy?.id || ''));
  $('restore-default').addEventListener('click', () => switchStrategyTo(BUILT_IN_STRATEGY_ID));
  $('delete-imported').addEventListener('click', requestDeleteImportedStrategy);
  $('cancel-delete-imported').addEventListener('click', hideDeleteConfirmation);
  $('confirm-delete-imported').addEventListener('click', deleteImportedStrategy);
  $('clear-template-upload').addEventListener('click', clearTemplateUpload);
  $('apply').addEventListener('click', applySelectedFixes);
  $('select-fixable').addEventListener('click', () => {
    if (operationBusy || !currentSummary) return;
    currentSummary.findings
      .filter(f => f.autoFixable && f.status === 'failed')
      .forEach(f => selectedRuleIds.add(f.ruleId));
    renderFindings();
    updateSelectedCount();
  });
  $('clear-selection').addEventListener('click', () => {
    if (operationBusy) return;
    selectedRuleIds.clear();
    renderFindings();
    updateSelectedCount();
  });

  document.querySelectorAll<HTMLButtonElement>('.filter').forEach(button => {
    button.addEventListener('click', () => {
      activeFilter = button.dataset.filter || 'all';
      document.querySelectorAll('.filter').forEach(b => b.classList.remove('active'));
      button.classList.add('active');
      renderFindings();
    });
  });
}

function strategyDisplayName(strategy: StrategyDefinition): string {
  if (strategy.origin === 'imported') return `Imported · ${strategy.sourceTemplateName || strategy.name}`;
  return 'AIP Default (Built-in)';
}

function refreshStrategyUi(): void {
  const available = StrategyStore.getAvailableStrategies();
  const select = $('strategy') as HTMLSelectElement;
  select.innerHTML = available.map(item => `<option value="${esc(item.id)}">${esc(strategyDisplayName(item))}</option>`).join('');
  select.value = activeStrategyId;
  $('rule-count').textContent = `${activeStrategy.rules.length} rules`;
  $('strategy-origin').textContent = activeStrategy.origin === 'imported'
    ? `Imported · based on ${activeStrategy.baseStrategyId || BUILT_IN_STRATEGY_ID}`
    : 'Built-in and immutable';

  const imported = available.find(item => item.origin === 'imported');
  $('template-management-status').textContent = imported
    ? `Current: ${strategyDisplayName(activeStrategy)} · ${imported.approvedChangeCount || 0} approved change(s) · Based on AIP Default`
    : `Current: ${strategyDisplayName(activeStrategy)}`;
  $('imported-strategy-actions').classList.toggle('hidden', !imported);
  if (!imported) hideDeleteConfirmation();
  updateOperationControlState();
}

function updateStrategyControlState(): void {
  const imported = StrategyStore.readState().importedStrategy;
  ($('strategy') as HTMLSelectElement).disabled = operationBusy;
  ($('upload-template') as HTMLButtonElement).disabled = operationBusy;
  ($('clear-template-upload') as HTMLButtonElement).disabled = operationBusy;
  ($('create-imported-strategy') as HTMLButtonElement).disabled =
    operationBusy || acceptedTemplateChanges.size === 0;
  ($('download-template-strategy') as HTMLButtonElement).disabled =
    operationBusy || acceptedTemplateChanges.size === 0;
  ($('use-imported-strategy') as HTMLButtonElement).disabled =
    operationBusy || !imported || activeStrategyId === imported?.id;
  ($('restore-default') as HTMLButtonElement).disabled =
    operationBusy || activeStrategyId === BUILT_IN_STRATEGY_ID;
  ($('delete-imported') as HTMLButtonElement).disabled = operationBusy || !imported;
  ($('cancel-delete-imported') as HTMLButtonElement).disabled = operationBusy;
  ($('confirm-delete-imported') as HTMLButtonElement).disabled = operationBusy || !imported;
  const safeIds = templateAnalysis ? safeTemplateProposalIds(templateAnalysis) : [];
  ($('select-all-template-changes') as HTMLButtonElement).disabled =
    operationBusy || !templateAnalysis || !safeIds.some(ruleId => !acceptedTemplateChanges.has(ruleId));
  ($('clear-template-changes') as HTMLButtonElement).disabled =
    operationBusy || acceptedTemplateChanges.size === 0;

  document
    .querySelectorAll<HTMLButtonElement>('button[data-template-action]')
    .forEach(button => {
      button.disabled = operationBusy;
    });
}

function updateOperationControlState(): void {
  ($('validate') as HTMLButtonElement).disabled = operationBusy || !wordHostAvailable;
  ($('apply') as HTMLButtonElement).disabled = operationBusy || !wordHostAvailable || selectedRuleIds.size === 0;
  ($('select-fixable') as HTMLButtonElement).disabled = operationBusy || !wordHostAvailable;
  ($('clear-selection') as HTMLButtonElement).disabled = operationBusy || !wordHostAvailable;

  document
    .querySelectorAll<HTMLInputElement>('input.approve')
    .forEach(input => {
      input.disabled = operationBusy || !wordHostAvailable;
    });

  document
    .querySelectorAll<HTMLButtonElement>('button[data-locate]')
    .forEach(button => {
      button.disabled = operationBusy || !wordHostAvailable;
    });

  updateStrategyControlState();
}

function clearValidationAfterStrategyChange(message: string): void {
  currentSummary = null;
  currentSnapshot = null;
  validatedDocumentFingerprint = null;
  validatedStrategyFingerprint = null;
  selectedRuleIds.clear();
  $('results').classList.add('hidden');
  updateSelectedCount();
  $('status-dot').textContent = message;
}

function switchStrategyTo(strategyId: string): void {
  if (!strategyId || operationBusy) return;
  try {
    StrategyStore.setActiveStrategy(strategyId);
    activeStrategyId = strategyId;
    activeStrategy = StrategyStore.getActiveStrategy();
    refreshStrategyUi();
    clearValidationAfterStrategyChange('Strategy changed. Please validate the document again.');
  } catch (error) {
    showError(error instanceof Error ? error.message : String(error));
  }
}

function handleStrategyChange(): void {
  switchStrategyTo(($('strategy') as HTMLSelectElement).value);
}

function hideDeleteConfirmation(): void {
  pendingDeleteStrategyId = null;
  $('delete-imported-confirmation').classList.add('hidden');
}

function requestDeleteImportedStrategy(): void {
  if (operationBusy) return;
  const imported = StrategyStore.readState().importedStrategy;
  if (!imported) return;
  pendingDeleteStrategyId = imported.id;
  $('delete-imported-name').textContent = imported.sourceTemplateName || imported.name;
  $('delete-imported-confirmation').classList.remove('hidden');
  updateOperationControlState();
}

function deleteImportedStrategy(): void {
  if (operationBusy || !pendingDeleteStrategyId) return;
  const strategyId = pendingDeleteStrategyId;
  const imported = StrategyStore.readState().importedStrategy;
  if (!imported || !isDeleteConfirmationCurrent(strategyId, imported.id)) {
    hideDeleteConfirmation();
    showError('The uploaded strategy changed. Please review it before deleting.');
    return;
  }
  const wasActive = activeStrategyId === strategyId;
  try {
    StrategyStore.removeImportedStrategy(strategyId);
    hideDeleteConfirmation();
    if (wasActive) {
      activeStrategyId = BUILT_IN_STRATEGY_ID;
      activeStrategy = StrategyStore.getBuiltInStrategy();
      refreshStrategyUi();
      clearValidationAfterStrategyChange('Uploaded strategy deleted. AIP Default is active. Please validate again.');
    } else {
      refreshStrategyUi();
      $('status-dot').textContent = 'Uploaded strategy deleted';
    }
  } catch (error) {
    hideDeleteConfirmation();
    showError(error instanceof Error ? error.message : String(error));
  }
}

function createImportedStrategy(): void {
  if (operationBusy || !templateAnalysis || acceptedTemplateChanges.size === 0) return;
  try {
    const imported = StrategyStore.createImportedStrategy(templateAnalysis, acceptedTemplateChanges);
    const wasImportedActive = activeStrategy.origin === 'imported';
    StrategyStore.saveImportedStrategy(imported);
    if (wasImportedActive) {
      StrategyStore.restoreBuiltInDefault();
      activeStrategyId = BUILT_IN_STRATEGY_ID;
      activeStrategy = StrategyStore.getBuiltInStrategy();
      clearValidationAfterStrategyChange('New imported strategy created. AIP Default is active; validate again.');
    }
    refreshStrategyUi();
    $('status-dot').textContent = wasImportedActive
      ? 'New imported strategy created. AIP Default is active; please validate again.'
      : 'Imported strategy created; AIP Default remains active';
  } catch (error) {
    showError(error instanceof Error ? error.message : String(error));
  }
}

async function analyzeUploadedTemplate(file: File): Promise<void> {
  if (operationBusy) return;
  templateAnalysis = null;
  acceptedTemplateChanges.clear();
  $('template-analysis').classList.add('hidden');
  setBusy(true, 'Analyzing the uploaded AIP template locally…');
  $('template-file-name').textContent = file.name;
  $('status-dot').textContent = 'Analyzing uploaded template locally';
  hideError();
  try {
    templateAnalysis = await TemplateAnalyzer.analyzeFile(file, builtInAipStrategy);
    renderTemplateAnalysis();
    $('status-dot').textContent = `Template analysis complete: ${file.name}`;
  } catch (error) {
    templateAnalysis = null;
    $('template-analysis').classList.add('hidden');
    showError(error instanceof Error ? error.message : String(error));
    $('status-dot').textContent = 'Template analysis error';
  } finally {
    ($('template-file') as HTMLInputElement).value = '';
    setBusy(false);
  }
}

function clearTemplateUpload(): void {
  if (operationBusy) return;
  templateAnalysis = null;
  acceptedTemplateChanges.clear();
  ($('template-file') as HTMLInputElement).value = '';
  $('template-file-name').textContent = 'No template selected';
  $('template-analysis').classList.add('hidden');
  updateTemplateSelectedCount();
  updateStrategyControlState();
  $('status-dot').textContent = 'Current template upload cleared';
}

function formatTemplateValue(value: unknown): string {
  if (value === undefined) return 'Not determinable';
  if (value && typeof value === 'object' && 'before' in value && 'after' in value) {
    const pair = value as { before: unknown; after: unknown };
    return `${String(pair.before)} / ${String(pair.after)} pt`;
  }
  if (typeof value === 'boolean') return value ? 'Yes' : 'No';
  return String(value);
}

function renderTemplateAnalysis(): void {
  if (!templateAnalysis) return;
  const analysis = templateAnalysis;
  $('template-analysis').classList.remove('hidden');
  $('template-summary').innerHTML = [
    ['Unchanged', analysis.counts.unchanged],
    ['Changes', analysis.counts.changed],
    ['Review', analysis.counts.review],
    ['New', analysis.counts.new],
    ['Not derivable', analysis.counts.notDeterminable]
  ].map(([label, count]) => `<div><strong>${esc(count)}</strong><span>${esc(label)}</span></div>`).join('');

  $('template-components').innerHTML = analysis.components.length
    ? analysis.components.map(component => `<span class="template-component">✓ ${esc(component.component)}${component.styleName ? ` · ${esc(component.styleName)}` : ''}</span>`).join('')
    : '<span class="template-file-name">No recognized template components found.</span>';

  updateTemplateSelectedCount();

  const visibleProposals = analysis.proposals.filter(proposal => proposal.status === 'changed' || proposal.status === 'new');
  const notDeterminable = analysis.proposals.filter(proposal => proposal.status === 'not-determinable');
  const notDeterminableDetails = notDeterminable.length
    ? `<details class="template-not-determinable"><summary>Not determinable (${notDeterminable.length})</summary>${notDeterminable.map(proposal => `<p><strong>${esc(proposal.ruleId)}</strong> · ${esc(proposal.note || 'This rule cannot be inferred reliably from DOCX formatting alone.')}</p>`).join('')}</details>`
    : '';
  $('template-proposals').innerHTML = analysis.warnings.length || visibleProposals.length || notDeterminable.length
    ? `${analysis.warnings.map(warning => `<p class="template-note">Warning: ${esc(warning)}</p>`).join('')}${visibleProposals.map(renderTemplateProposal).join('')}${notDeterminableDetails}`
    : '<p class="template-note">No measurable strategy changes detected. Existing strategy values are unchanged.</p>';

  updateStrategyControlState();
  $('template-proposals').querySelectorAll<HTMLButtonElement>('button[data-template-action]').forEach(button => {
    button.addEventListener('click', () => {
      if (operationBusy) return;
      const ruleId = button.dataset.templateRuleId;
      if (!ruleId) return;
      if (button.dataset.templateAction === 'accept') acceptedTemplateChanges.add(ruleId);
      else acceptedTemplateChanges.delete(ruleId);
      renderTemplateAnalysis();
    });
  });
}

function updateTemplateSelectedCount(): void {
  const safeCount = templateAnalysis
    ? safeTemplateProposalIds(templateAnalysis).length
    : 0;
  const selectedSafeCount = templateAnalysis
    ? safeTemplateProposalIds(templateAnalysis)
      .filter(ruleId => acceptedTemplateChanges.has(ruleId)).length
    : 0;
  $('template-selected-count').textContent =
    `${selectedSafeCount} of ${safeCount} safe changes selected`;
}

function renderTemplateProposal(proposal: StrategyChangeProposal): string {
  const accepted = acceptedTemplateChanges.has(proposal.ruleId);
  const canAccept = isTemplateProposalAcceptable(proposal);
  const action = canAccept
    ? `<button data-template-action="${accepted ? 'keep' : 'accept'}" data-template-rule-id="${esc(proposal.ruleId)}">${accepted ? 'Keep current' : 'Accept proposal'}</button>`
    : '';
  return `<article class="template-proposal ${esc(proposal.status)}">
    <div class="template-proposal-title"><span>${esc(proposal.ruleId)} · ${esc(proposal.ruleName)}</span><span>${esc(proposal.confidence)} ${accepted ? '· accepted' : ''}</span></div>
    <p>${esc(proposal.component)} · ${esc(proposal.property)}</p>
    <div class="compare"><div><span>Current strategy</span><strong>${esc(formatTemplateValue(proposal.currentValue))}</strong></div><div><span>Uploaded template</span><strong>${esc(formatTemplateValue(proposal.detectedValue))}</strong></div></div>
    ${proposal.note ? `<p>${esc(proposal.note)}</p>` : ''}
    <div class="template-proposal-actions">${action}</div>
  </article>`;
}

function downloadApprovedTemplateStrategy(): void {
  if (operationBusy || !templateAnalysis || !acceptedTemplateChanges.size) return;
  const reviewed = TemplateAnalyzer.applyAcceptedChanges(builtInAipStrategy, templateAnalysis, acceptedTemplateChanges);
  const blob = new Blob([JSON.stringify(reviewed, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${templateAnalysis.fileName.replace(/\.docx$/i, '')}-reviewed-strategy.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  $('status-dot').textContent = 'Downloaded reviewed strategy copy';
}

async function validateDocument(): Promise<boolean> {
  setBusy(true, 'Reading the open Word document and inspecting DOCX/OOXML…');
  selectedRuleIds.clear();
  validatedDocumentFingerprint = null;
  validatedStrategyFingerprint = null;
  currentSnapshot = null;
  currentSummary = null;
  $('results').classList.add('hidden');
  updateSelectedCount();
  hideError();

  try {
    const snapshot = await WordDocumentAdapter.snapshot();
    currentSnapshot = snapshot;
    // The current UI exposes manuscript validation only. Submission/readiness
    // rules remain available in the strategy for a future dedicated flow.
    if (snapshot.features) snapshot.features.submissionMode = false;
    activeStrategy = StrategyStore.getActiveStrategy();
    currentSummary = ValidationEngine.validate(activeStrategy, snapshot);
    validatedDocumentFingerprint = documentFingerprint(snapshot);
    validatedStrategyFingerprint = strategyFingerprint(activeStrategy);
    renderSummary();
    $('results').classList.remove('hidden');
    $('status-dot').textContent = snapshot.layout.ooxmlAvailable
      ? `Validated with Office.js + OOXML (${snapshot.layout.ooxmlParser || 'parsed'})`
      : 'Validated with Office.js';
    return true;
  } catch (error) {
    showError(error instanceof Error ? error.message : String(error));
    $('status-dot').textContent = 'Validation error';
    return false;
  } finally {
    setBusy(false);
  }
}

function renderSummary(): void {
  if (!currentSummary) return;

  const s = currentSummary;
  $('active-strategy-summary').textContent = `Strategy: ${strategyDisplayName(activeStrategy)}`;
  $('score').textContent = `${s.overallScore}%`;
  $('score-ring').style.setProperty('--score', String(Math.max(0, Math.min(100, s.overallScore))));
  const issueSummary = s.unresolvedFailures === 0
    ? 'No deterministic formatting issues detected'
    : `${s.unresolvedFailures} issue${s.unresolvedFailures === 1 ? '' : 's'} remaining · ${s.fixableFailures} safe fix${s.fixableFailures === 1 ? '' : 'es'} available`;
  $('score-detail').textContent = `AIP compliance: ${s.complianceScore}% · ${issueSummary}`;
  const technical = s.notEvaluated ? ` · ${s.notEvaluated} not evaluated` : '';
  // Keep each counter identical to the findings shown by its tab filter.
  $('stat-all').textContent = String(filterFindings(s.findings, 'all').length);
  $('coverage').textContent = `${s.scoredRules} compliance scored · ${s.consistencyPassed + s.consistencyFailed} consistency checks · ${s.manualReview} review · ${s.checklist} checklist · ${s.notApplicable} not applicable${technical}`;
  // Consistency remains excluded from the compliance score itself, while its
  // findings remain visible in the status and consistency tabs.
  $('stat-failed').textContent = String(filterFindings(s.findings, 'failed').length);
  $('stat-manual').textContent = String(s.manualReview);
  $('stat-checklist').textContent = String(s.checklist);
  $('stat-passed').textContent = String(filterFindings(s.findings, 'passed').length);
  $('stat-consistency').textContent = String(filterFindings(s.findings, 'consistency').length);

  const hasConflicts = s.findings.some(f => f.reviewKind === 'source_conflict');
  $('conflict-banner').classList.toggle('hidden', !hasConflicts);

  renderFindings();
  updateSelectedCount();
}

function renderFindings(): void {
  if (!currentSummary) return;

  const findings = filterFindings(currentSummary.findings, activeFilter);
  $('visible-count').textContent = `${findings.length} shown`;

  const container = $('findings');
  container.innerHTML = findings.map(renderFinding).join('');

  container.querySelectorAll<HTMLInputElement>('input[data-rule-id]').forEach(box => {
    box.addEventListener('change', () => {
      const id = box.dataset.ruleId!;
      box.checked ? selectedRuleIds.add(id) : selectedRuleIds.delete(id);
      updateSelectedCount();
    });
  });

  container.querySelectorAll<HTMLButtonElement>('button[data-locate]').forEach(button => {
    button.addEventListener('click', async () => {
      if (operationBusy || !wordHostAvailable) return;
      const index = Number(button.dataset.locate);
      try {
        const expectedText = currentSnapshot?.paragraphs.find(paragraph => paragraph.index === index)?.text;
        await WordDocumentAdapter.locateParagraph(index, expectedText);
      } catch (error) {
        showError(error instanceof Error ? error.message : String(error));
      }
    });
  });
}

function renderFinding(f: Finding): string {
  const statusLabel = f.assessmentRole === 'consistency' && f.status === 'failed'
    ? 'consistency issue'
    : f.status === 'checklist' ? 'submission checklist' : f.status.replace(/_/g, ' ');
  const checked = selectedRuleIds.has(f.ruleId) ? 'checked' : '';
  const checkbox = f.autoFixable && f.status === 'failed'
    ? `<input class="approve" type="checkbox" data-rule-id="${esc(f.ruleId)}" ${checked} ${operationBusy ? 'disabled' : ''} aria-label="Approve ${esc(f.ruleName)} fix">`
    : '<span></span>';

  const locateIndex = f.paragraphIndices[0];
  const locateText = locateIndex === undefined
    ? undefined
    : currentSnapshot?.paragraphs.find(paragraph => paragraph.index === locateIndex)?.text;
  const locate = locateIndex !== undefined
    ? `<button data-locate="${locateIndex}" data-locate-fingerprint="${encodeURIComponent(paragraphTextFingerprint(locateText || ''))}" ${operationBusy || !wordHostAvailable ? 'disabled' : ''}>Locate</button>`
    : '';

  const fixBadge = f.autoFixable
    ? `<span class="badge">${esc(f.target === 'document' ? 'structural fix' : f.property.startsWith('text.') ? 'text fix' : 'safe fix')}</span>`
    : '';
  const reviewBadge = f.reviewKind === 'manual'
    ? '<span class="badge">human judgement</span>'
    : f.reviewKind === 'submission'
      ? '<span class="badge">outside DOCX</span>'
      : f.reviewKind === 'technical'
        ? '<span class="badge">technical limitation</span>'
        : f.reviewKind === 'source_conflict'
          ? '<span class="badge">source conflict</span>'
          : f.assessmentRole === 'consistency'
            ? '<span class="badge">consistency check</span>'
          : '';

  return `
    <article class="finding ${esc(f.status)}">
      <div class="finding-main">
        ${checkbox}
        <div>
          <span class="badge status-${esc(f.status)}">${esc(statusLabel)}</span>
          <span class="badge">${esc(f.category)}</span>
          ${fixBadge}
          ${reviewBadge}
          <h3>${esc(f.ruleId)} · ${esc(f.ruleName)}</h3>
          <p>${esc(f.message)}</p>
          ${f.paragraphIndices.length ? `<p class="affected"><strong>Affected paragraphs:</strong> ${esc(f.paragraphIndices.map(index => index + 1).join(', '))}</p>` : ''}
          <div class="compare">
            <div><span>Expected</span><strong>${esc(f.expected)}</strong></div>
            <div><span>Detected</span><strong>${esc(f.detected)}</strong></div>
          </div>
        </div>
        <div class="finding-actions">${locate}</div>
      </div>
      <details>
        <summary>Rule evidence</summary>
        <p>${esc(f.evidence)}</p>
        ${f.note ? `<p><strong>Note:</strong> ${esc(f.note)}</p>` : ''}
      </details>
    </article>`;
}

function updateSelectedCount(): void {
  const count = selectedRuleIds.size;
  $('selected-count').textContent = `${count} fix${count === 1 ? '' : 'es'} selected`;
  updateOperationControlState();
}

async function applySelectedFixes(): Promise<void> {
  const summary = currentSummary;
  if (!canStartCorrection(operationBusy, Boolean(summary), selectedRuleIds.size) || !summary) return;
  const selected = summary.findings.filter(f => selectedRuleIds.has(f.ruleId) && f.autoFixable && f.status === 'failed');
  if (!selected.length) return;
  const requestedRuleIds = selected.map(f => f.ruleId);

  setBusy(true, `Applying ${selected.length} approved deterministic fix${selected.length === 1 ? '' : 'es'}…`);
  hideError();

  try {
    const currentSnapshot = await WordDocumentAdapter.snapshot();
    if (
      !validatedDocumentFingerprint
      || documentFingerprint(currentSnapshot) !== validatedDocumentFingerprint
      || !validatedStrategyFingerprint
      || strategyFingerprint(StrategyStore.getActiveStrategy()) !== validatedStrategyFingerprint
    ) {
      throw new Error('The document or active strategy changed since validation. Please revalidate before applying safe fixes.');
    }
    let correctionResult: CorrectionResult;
    let correctionError: string | null = null;
    try {
      correctionResult = await CorrectionService.apply(selected, activeStrategy);
    } catch (error) {
      correctionError = error instanceof Error ? error.message : String(error);
      correctionResult = {
        verifications: [],
        attempts: [],
        requestedFixes: selected.length,
        executedFixes: 0,
        executionFailures: selected.length
      };
    }
    const validationSucceeded = await validateDocument();
    const findingByRuleId = new Map(
      currentSummary?.findings.map(finding => [finding.ruleId, finding]) || []
    );
    // A correction is resolved only when the refreshed validator explicitly
    // passes the same applicable rule. Missing evidence and not_applicable
    // are intentionally unresolved because neither proves compliance.
    const unresolved = requestedRuleIds.filter(ruleId =>
      !validationSucceeded
      || findingByRuleId.get(ruleId)?.status !== 'passed'
    );
    const resolvedCount = requestedRuleIds.length - unresolved.length;
    correctionResult.resolvedFixes = resolvedCount;
    correctionResult.unresolvedFixes = unresolved.length;
    correctionResult.attempts.forEach(attempt => {
      const finding = findingByRuleId.get(attempt.ruleId);
      attempt.resolvedAfterValidation = validationSucceeded
        && finding?.status === 'passed';
    });
    const executionNotes = [
      correctionResult.executionFailures > 0
        ? `${correctionResult.executionFailures} approved fix(es) failed during execution.`
        : '',
      correctionError ? `Correction batch error: ${correctionError}` : '',
      ...correctionResult.attempts
        .map(attempt => attempt.message)
        .filter((message): message is string => Boolean(message))
        .map(message => `Correction detail: ${message}`)
    ].filter(Boolean);
    const executionNote = executionNotes.length ? ` ${executionNotes.join(' ')}` : '';
    const unresolvedNote = unresolved.length
      ? ` Still requiring attention: ${unresolved.join(', ')}.`
      : '';
    $('status-dot').textContent = `Last correction batch: ${correctionResult.requestedFixes} approved fix${correctionResult.requestedFixes === 1 ? '' : 'es'} requested; ${correctionResult.executedFixes} executed; ${resolvedCount} resolved; ${unresolved.length} still require attention.${executionNote}${unresolvedNote}`;
  } catch (error) {
    showError(error instanceof Error ? error.message : String(error));
    $('status-dot').textContent = 'Correction error';
  } finally {
    setBusy(false);
  }
}

function setBusy(busy: boolean, message?: string): void {
  busyDepth = nextBusyDepth(busyDepth, busy);
  operationBusy = busyDepth > 0;
  updateOperationControlState();
  $('progress').classList.toggle('hidden', !operationBusy);
  if (message) $('progress').querySelector('p')!.textContent = message;
}

function showError(message: string): void {
  $('error').textContent = message;
  $('error').classList.remove('hidden');
}

function hideError(): void {
  $('error').classList.add('hidden');
  $('error').textContent = '';
}
