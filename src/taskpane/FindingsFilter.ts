import { Finding } from '../core/models';

export function filterFindings(findings: Finding[], activeFilter: string): Finding[] {
  return findings.filter(finding => {
    if (activeFilter === 'all') {
      return finding.status !== 'not_applicable';
    }

    if (activeFilter === 'failed') {
      // Failed consistency checks remain visible here so users can discover
      // and approve their safe fixes. Their assessmentRole still keeps them
      // out of the compliance score and the Consistency view remains available.
      return finding.status === 'failed';
    }

    if (activeFilter === 'consistency') {
      return finding.assessmentRole === 'consistency';
    }

    return finding.status === activeFilter;
  });
}
