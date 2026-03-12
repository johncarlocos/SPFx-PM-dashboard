import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IProject, IRfi } from '../models/IProject';

const LIST_PROJ = '3Edge_Projects';
const LIST_RFI = '3Edge_RFIs';

export class SharePointService {
  private _siteUrl: string;
  private _http: SPHttpClient;

  constructor(siteUrl: string, spHttpClient: SPHttpClient) {
    this._siteUrl = siteUrl;
    this._http = spHttpClient;
  }

  private async spGet(path: string): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const r: SPHttpClientResponse = await this._http.get(
      this._siteUrl + path,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); msg = e.error?.message?.value || msg; } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
    return r.json();
  }

  private async spPost(path: string, body: any): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const opts: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata'
      },
      body: JSON.stringify(body)
    };
    const r: SPHttpClientResponse = await this._http.post(
      this._siteUrl + path,
      SPHttpClient.configurations.v1,
      opts
    );
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); msg = e.error?.message?.value || msg; } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
    const text = await r.text();
    return text ? JSON.parse(text) : {};
  }

  private async spMerge(path: string, body: any): Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const opts: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify(body)
    };
    const r: SPHttpClientResponse = await this._http.post(
      this._siteUrl + path,
      SPHttpClient.configurations.v1,
      opts
    );
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); msg = e.error?.message?.value || msg; } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
  }

  private async spDelete(path: string): Promise<void> {
    const opts: ISPHttpClientOptions = {
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    };
    const r: SPHttpClientResponse = await this._http.post(
      this._siteUrl + path,
      SPHttpClient.configurations.v1,
      opts
    );
    if (!r.ok && r.status !== 404) throw new Error('DELETE HTTP ' + r.status);
  }

  // ── Project CRUD ──────────────────────────────────────────

  public async loadProjects(): Promise<IProject[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items?$top=500&$orderby=projNum asc`);
    return (d.value || []).map((i: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
      id: i.projNum || String(i.Id),
      spId: i.Id,
      projNum: i.projNum || '',
      name: i.name || i.Title || '',
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
      startDate: i.startDate ? i.startDate.substring(0, 10) : '',
      finishDate: i.finishDate ? i.finishDate.substring(0, 10) : '',
      ifaDate: i.ifaDate ? i.ifaDate.substring(0, 10) : '',
      ifcDate: i.ifcDate ? i.ifcDate.substring(0, 10) : '',
      detailers: i.detailers || '',
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
      startDate: d.startDate || null,
      finishDate: d.finishDate || null,
      ifaDate: d.ifaDate || null,
      ifcDate: d.ifcDate || null,
      detailers: d.detailers || '',
      isEwo: d.isEwo || false,
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
      dateIssued: i.dateIssued ? i.dateIssued.substring(0, 10) : '',
      dateRequired: i.dateRequired ? i.dateRequired.substring(0, 10) : '',
      description: i.description || '',
      attachments: i.attachments || '',
      clientRfi: i.clientRfi || '',
      dateReceived: i.dateReceived ? i.dateReceived.substring(0, 10) : '',
      response: i.response || '',
      responseDesc: i.responseDesc || '',
      sentBy: i.sentBy || '',
      sentByCompany: i.sentByCompany || '',
      impacted: i.impacted || 'No',
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
      attachments: d.attachments || '',
      clientRfi: d.clientRfi || '',
      dateReceived: d.dateReceived || null,
      response: d.response || '',
      responseDesc: d.responseDesc || '',
      sentBy: d.sentBy || '',
      sentByCompany: d.sentByCompany || '',
      impacted: d.impacted || 'No',
      ewoRef: d.ewoRef || '',
      ewoCcn: d.ewoCcn || '',
      tracked: d.tracked || false,
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
}
