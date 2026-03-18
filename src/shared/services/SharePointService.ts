import { SPHttpClient } from '@microsoft/sp-http';
import { IProject, IRfi } from '../models/IProject';

const LIST_PROJ = '3Edge_Projects';
const LIST_RFI = '3Edge_RFIs';

export class SharePointService {
  private _siteUrl: string;
  private _digest: string = '';

  // spHttpClient param kept for API compat but all requests use plain fetch
  constructor(siteUrl: string, _spHttpClient?: SPHttpClient) {
    this._siteUrl = siteUrl;
  }

  private parseDate(val: string | null | undefined): string {
    if (!val) return '';
    const m = val.match(/\/Date\((\d+)\)\//);
    if (m) return new Date(Number(m[1])).toISOString().substring(0, 10);
    return val.substring(0, 10);
  }

  private async getDigest(): Promise<string> {
    if (this._digest) return this._digest;
    const r = await fetch(this._siteUrl + '/_api/contextinfo', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Accept': 'application/json;odata=nometadata' }
    });
    if (r.ok) {
      const data = await r.json();
      this._digest = data.FormDigestValue || '';
    }
    return this._digest;
  }

  private async spGet(path: string): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const r = await fetch(this._siteUrl + path, {
      credentials: 'include',
      headers: { Accept: 'application/json;odata=nometadata' }
    });
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); const em = e.error?.message; msg = (typeof em === 'object' && em?.value) ? em.value : (em || e.error?.code || msg); } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
    return r.json();
  }

  private async spPost(path: string, body: any): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const digest = await this.getDigest();
    const r = await fetch(this._siteUrl + path, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      },
      body: JSON.stringify(body)
    });
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); const em = e.error?.message; msg = (typeof em === 'object' && em?.value) ? em.value : (em || e.error?.code || msg); } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
    const text = await r.text();
    return text ? JSON.parse(text) : {};
  }

  private async spMerge(path: string, body: any): Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const digest = await this.getDigest();
    const r = await fetch(this._siteUrl + path, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        'X-RequestDigest': digest
      },
      body: JSON.stringify(body)
    });
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); const em = e.error?.message; msg = (typeof em === 'object' && em?.value) ? em.value : (em || e.error?.code || msg); } catch (_x) { /* ignore */ }
      // if digest expired, clear cache so next call refreshes it
      if (r.status === 403) this._digest = '';
      throw new Error(msg);
    }
  }

  private async spDelete(path: string): Promise<void> {
    const digest = await this.getDigest();
    const r = await fetch(this._siteUrl + path, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
        'X-RequestDigest': digest
      }
    });
    if (!r.ok && r.status !== 404) {
      if (r.status === 403) this._digest = '';
      throw new Error('DELETE HTTP ' + r.status);
    }
  }

  // ── Project CRUD ──────────────────────────────────────────

  public async loadProjects(): Promise<IProject[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items?$top=500&$orderby=projNum asc`);
    return (d.value || []).map((i: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
      id: i.projNum || String(i.Id),
      spId: i.Id,
      projNum: i.projNum || '',
      name: i.name || i.Title || '',
      discipline: i.discipline || '',
      status: i.status || 'Active',
      year: i.year ? Number(i.year) : new Date().getFullYear(),
      hrsAllowed: Number(i.hrsAllowed) || 0,
      hrsUsed: Number(i.hrsUsed) || 0,
      rfisAllowed: Number(i.rfisAllowed) || 0,
      quoteNum: i.quoteNum || '',
      contact: i.contact || '',
      company: i.company || '',
      email: i.email || '',
      mobile: i.mobile || '',
      clientNum: i.clientNum || '',
      clientp0: i.clientp0 || '',
      startDate: this.parseDate(i.startDate),
      finishDate: this.parseDate(i.finishDate),
      ifaDate: this.parseDate(i.ifaDate),
      ifcDate: this.parseDate(i.ifcDate),
      detailers: i.detailers || '',
      teamLead: i.teamLead || '',
      teamMembers: i.teamMembers || '',
      notes: i.notes || '',
      isEwo: i.isEwo || false,
      ewoNum: i.ewoNum || '',
      parentId: i.parentId || null
    }));
  }

  private pBody(d: IProject): object {
    return {
      Title: d.projNum || '',
      projNum: d.projNum || '',
      name: d.name || '',
      discipline: d.discipline || '',
      status: d.status || 'Active',
      year: Number(d.year) || new Date().getFullYear(),
      hrsAllowed: Number(d.hrsAllowed) || 0,
      hrsUsed: Number(d.hrsUsed) || 0,
      rfisAllowed: Number(d.rfisAllowed) || 0,
      quoteNum: d.quoteNum || '',
      contact: d.contact || '',
      company: d.company || '',
      email: d.email || '',
      mobile: d.mobile || '',
      clientNum: d.clientNum || '',
      clientp0: d.clientp0 || '',
      startDate: d.startDate || null,
      finishDate: d.finishDate || null,
      ifaDate: d.ifaDate || null,
      ifcDate: d.ifcDate || null,
      detailers: d.detailers || '',
      teamLead: d.teamLead || '',
      teamMembers: d.teamMembers || '',
      notes: d.notes || '',
      isEwo: d.isEwo === true || (d.isEwo as unknown as string) === 'true',
      ewoNum: d.ewoNum || '',
      parentId: d.parentId || null
    };
  }

  public async addProject(d: IProject): Promise<number> {
    const r = await this.spPost(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items`, this.pBody(d));
    if (!r || !r.Id) throw new Error('No Id returned');
    return r.Id;
  }

  public async updateProject(spId: number, d: IProject): Promise<void> {
    await this.spMerge(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items(${spId})`, this.pBody(d));
  }

  public async deleteProject(spId: number): Promise<void> {
    await this.spDelete(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items(${spId})`);
  }

  // ── Email ─────────────────────────────────────────────────

  public async sendEmail(to: string, cc: string, subject: string, body: string): Promise<void> {
    const digest = await this.getDigest();
    const toList = to.split(/[,;]/).map(s => s.trim()).filter(Boolean);
    const ccList = cc ? cc.split(/[,;]/).map(s => s.trim()).filter(Boolean) : [];
    const r = await fetch(this._siteUrl + '/_api/SP.Utilities.Utility.SendEmail', {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      body: JSON.stringify({
        properties: {
          '__metadata': { type: 'SP.Utilities.EmailProperties' },
          'To': { results: toList },
          'CC': { results: ccList },
          'Subject': subject,
          'Body': body
        }
      })
    });
    if (!r.ok) {
      let msg = 'Email failed: HTTP ' + r.status;
      try { const e = await r.json(); msg = e.error?.message || msg; } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
  }

  // ── RFI CRUD ──────────────────────────────────────────────

  public async loadRfis(): Promise<IRfi[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_RFI}')/items?$top=1000&$orderby=rfiSeq asc`);
    return (d.value || []).map((i: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
      id: i.rfiNum || String(i.Id),
      spId: i.Id,
      rfiNum: i.rfiNum || '',
      rfiSeq: Number(i.rfiSeq) || 0,
      projectId: i.projectId || '',
      projectName: i.projectName || '',
      rfiType: i.rfiType || '',
      status: i.status || 'Open',
      submittedTo: i.submittedTo || '',
      toCompany: i.toCompany || '',
      by: i.by || '',
      byCompany: i.byCompany || '',
      cc: i.cc || '',
      dateIssued: this.parseDate(i.dateIssued),
      dateRequired: this.parseDate(i.dateRequired),
      description: i.description || '',
      attachments: i.attachments || '',
      clientRfi: i.clientRfi || '',
      dateReceived: this.parseDate(i.dateReceived),
      response: i.response || '',
      responseDesc: i.responseDesc || '',
      sentBy: i.sentBy || '',
      sentByCompany: i.sentByCompany || '',
      impacted: i.impacted === true ? 'Yes' : (i.impacted === false ? 'No' : (i.impacted || 'No')),
      ewoRef: i.ewoRef || '',
      ewoCcn: i.ewoCcn || '',
      tracked: i.tracked || false,
      model: Number(i.model) || 0,
      connections: Number(i.connections) || 0,
      checking: Number(i.checking) || 0,
      drawings: Number(i.drawings) || 0,
      admin: Number(i.admin) || 0,
      revision: i.revision || 'A',
      email: i.email || ''
    }));
  }

  private rBody(d: IRfi): object {
    return {
      Title: d.rfiNum || '',
      rfiNum: d.rfiNum || '',
      rfiSeq: Number(d.rfiSeq) || 0,
      projectId: d.projectId || '',
      projectName: d.projectName || '',
      rfiType: d.rfiType || '',
      status: d.status || 'Open',
      submittedTo: d.submittedTo || '',
      toCompany: d.toCompany || '',
      by: d.by || '',
      byCompany: d.byCompany || '',
      cc: d.cc || '',
      dateIssued: d.dateIssued || null,
      dateRequired: d.dateRequired || null,
      description: d.description || '',
      clientRfi: d.clientRfi || '',
      dateReceived: d.dateReceived || null,
      response: d.response || '',
      responseDesc: d.responseDesc || '',
      sentBy: d.sentBy || '',
      sentByCompany: d.sentByCompany || '',
      impacted: d.impacted === 'Yes',
      ewoRef: d.ewoRef || '',
      ewoCcn: d.ewoCcn || '',
      model: Number(d.model) || 0,
      connections: Number(d.connections) || 0,
      checking: Number(d.checking) || 0,
      drawings: Number(d.drawings) || 0,
      admin: Number(d.admin) || 0,
      revision: d.revision || 'A',
      email: d.email || ''
    };
  }

  public async addRfi(d: IRfi): Promise<number> {
    const r = await this.spPost(`/_api/web/lists/getbytitle('${LIST_RFI}')/items`, this.rBody(d));
    if (!r || !r.Id) throw new Error('No Id returned');
    return r.Id;
  }

  public async updateRfi(spId: number, d: IRfi): Promise<void> {
    await this.spMerge(`/_api/web/lists/getbytitle('${LIST_RFI}')/items(${spId})`, this.rBody(d));
  }

  public async deleteRfi(spId: number): Promise<void> {
    await this.spDelete(`/_api/web/lists/getbytitle('${LIST_RFI}')/items(${spId})`);
  }

  // ── Attachments ──────────────────────────────────────────────

  public async getAttachments(spId: number): Promise<{ FileName: string; ServerRelativeUrl: string }[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_RFI}')/items(${spId})/AttachmentFiles`);
    return d.value || [];
  }

  public async uploadAttachment(spId: number, file: File): Promise<void> {
    const digest = await this.getDigest();
    const buf = await file.arrayBuffer();
    const url = this._siteUrl +
      `/_api/web/lists/getbytitle('${LIST_RFI}')/items(${spId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
    const r = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      },
      body: buf
    });
    if (!r.ok) {
      if (r.status === 403) this._digest = '';
      let msg = 'Upload failed: HTTP ' + r.status;
      try { const e = await r.json(); msg = e.error?.message || msg; } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
  }

  public async deleteAttachment(spId: number, fileName: string): Promise<void> {
    await this.spDelete(
      `/_api/web/lists/getbytitle('${LIST_RFI}')/items(${spId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`
    );
  }
}
