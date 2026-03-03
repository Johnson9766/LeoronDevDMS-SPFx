import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import homeHTML from './homePageHtml';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'WpHomePageWebPartStrings';

export interface IWpHomePageWebPartProps {
  description: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface TreeNode {
  id: string;
  label: string;
  type?: 'link';
  iconType?: 'FolderIcon' | 'linkIcon';
  url?: string;
  target?: string;
  children: TreeNode[];
}

/**
 * Dynamic payload — one key per document library.
 *
 * Key format:  "tabGridData_<slug>"
 * e.g.:
 *   {
 *     tabGridData_department_functions: [ ...TreeNode[] ],
 *     tabGridData_hr:                   [ ...TreeNode[] ],
 *   }
 *
 * Each key matches the data-tab-obj on the injected HTML tab div.
 * home.js spreads this into TAB_GRID_DATA and looks up by that key.
 */
export interface TabGridPayload {
  [key: string]: TreeNode[];
}

interface SPList {
  Id: string;
  Title: string;
  BaseTemplate: number;
  Hidden: boolean;
  RootFolder: { ServerRelativeUrl: string };
  DefaultView?: { Id: string };
}

interface SPFolder {
  UniqueId: string;
  Name: string;
  ServerRelativeUrl: string;
}

interface SPFile {
  UniqueId: string;
  Name: string;
  ServerRelativeUrl: string;
  LinkingUrl: string;
}

interface SPBannerItem {
  Id: number;
  FileRef: string;        // server-relative URL of the image file
  FileLeafRef: string;    // filename e.g. "banner.png"
  TimeCreated: string;
  ActiveStatus: boolean;  // Yes = true, No = false
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const EXCLUDED_LIBRARIES: string[] = [
  'Form Templates',
  'Style Library',
  'Site Assets',
  'Site Pages',
  'Preservation Hold Library',
  'Pages',
  'Images',
  'Documents',
  'Site Banners'
];

const MAX_DEPTH = 4;

/**
 * Converts a library title to a TAB_GRID_DATA key.
 * "Department Functions" → "tabGridData_department_functions"
 * Must match the data-tab-obj value injected into the HTML tab div.
 */
function toTabKey(title: string): string {
  const slug = title
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_|_$/g, '');
  return `tabGridData_${slug}`;
}

/**
 * Slug for folder node IDs used in the DOM.
 * "Sales Team" → "sales-team"
 */
function toId(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
}

// ─────────────────────────────────────────────────────────────────────────────
// WebPart
// ─────────────────────────────────────────────────────────────────────────────

export default class WpHomePageWebPart extends BaseClientSideWebPart<IWpHomePageWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = homeHTML.allElementsHtml;

    const workbenchContent = document.getElementById('workbenchPageContent');
    if (workbenchContent) {
      workbenchContent.style.maxWidth = 'none';
    }

    this.loadHome();
  }

  /**
   * Full flow:
   * 1. Fetch all non-default doc libraries from SharePoint
   * 2. Build a TreeNode[] for each (first-level folders = root nodes)
   * 3. Inject <div data-tab-obj="tabGridData_<slug>"> tabs into the HTML
   * 4. Set window["TAB_GRID_DATA"] = { tabGridData_<slug>: [...], ... }
   * 5. Load home.js — it reads TAB_GRID_DATA and renders the active tab
   */
 private loadHome(): void {
  const baseUrl = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/resources`;

  Promise.all([
    this._getTabGridData(),
    this._getBannerImages(),           // ← NEW
  ])
    .then(([result, banners]) => {
      const { payload, libraries } = result;

      // ── 1. Inject banner slides ──────────────────────────────────────
      this._renderBanners(banners);    // ← NEW

      // ── 2. Inject dynamic tab divs BEFORE home.js loads ─────────────
      this._renderTabs(libraries);

      // ── 3. Set global payload for home.js ────────────────────────────
      (window as unknown as Record<string, unknown>)['TAB_GRID_DATA'] = payload;

      // ── 4. Load scripts ───────────────────────────────────────────────
      return SPComponentLoader.loadScript(`${baseUrl}/js/swiper-bundle.min.js`)
        .then(() => SPComponentLoader.loadScript(`${baseUrl}/js/home.js`));
    })
    .catch((err) => {
      console.error('[DMS] Failed to load data:', err);
      SPComponentLoader.loadScript(`${baseUrl}/js/home.js`).catch(
        (scriptErr) => console.error('[DMS] home.js load error:', scriptErr)
      );
    });
}


  //Banner Image start
private async _getBannerImages(): Promise<SPBannerItem[]> {
  const siteUrl = this.context.pageContext.web.absoluteUrl;
  const libraryName = encodeURIComponent('Site Banners'); // ← match your library name exactly


  const endpoint =
    `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items` +
    `?$select=Id,FileRef,FileLeafRef,Created,ActiveStatus` +
    `&$filter=ActiveStatus eq 1 and FSObjType eq 0` +
    `&$orderby=Created desc` +
    `&$top=2`;

  try {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      console.warn(`[DMS] Banner fetch failed: ${response.status} ${response.statusText}`);
      return [];
    }

    const data = await response.json();
    return (data.value as SPBannerItem[]) || [];

  } catch (error) {
    console.warn('[DMS] Error fetching banners:', error);
    return [];
  }
}

private _renderBanners(banners: SPBannerItem[]): void {
  const swiperWrapper = this.domElement.querySelector('.swiper-wrapper');
  const bannerWrapper = this.domElement.querySelector('.leoron-banner-wrapper') as HTMLElement;

  if (!swiperWrapper) {
    console.warn('[DMS] .swiper-wrapper not found.');
    return;
  }

  if (banners.length === 0) {
    // Hide the entire banner section including title overlay
    if (bannerWrapper) {
      bannerWrapper.style.display = 'none';
    }
    console.warn('[DMS] No active banners — banner section hidden.');
    return;
  }

  const origin = new URL(this.context.pageContext.web.absoluteUrl).origin;

  // Clear static placeholder slides
  swiperWrapper.innerHTML = '';

  banners.forEach((item) => {
    const imageUrl = `${origin}${item.FileRef}`;

    const slide = document.createElement('div');
    slide.className = 'leoron-banner-swiper-item swiper-slide';

    const img = document.createElement('img');
    img.src = imageUrl;
    img.alt = item.FileLeafRef.replace(/\.[^/.]+$/, '');

    slide.appendChild(img);
    swiperWrapper.appendChild(slide);
  });
}

  /**
   * Clears the static placeholder tabs from the HTML template and
   * injects one real tab per document library.
   *
   * Output HTML example (3 libraries):
   *
   *   <div data-tab-obj="tabGridData_department_functions"
   *        class="home-page-tab-title flex-fill text-center home-page-tab-title-active">
   *     Department Functions
   *   </div>
   *   <div data-tab-obj="tabGridData_centralized_functions"
   *        class="home-page-tab-title flex-fill text-center">
   *     Centralized Functions
   *   </div>
   *   <div data-tab-obj="tabGridData_hr"
   *        class="home-page-tab-title flex-fill text-center">
   *     HR
   *   </div>
   *
   * - data-tab-obj value MUST match the key in window["TAB_GRID_DATA"]
   * - First tab gets "home-page-tab-title-active" — home.js renders it on load
   */
  private _renderTabs(libraries: SPList[]): void {
    const wrapper = this.domElement.querySelector('.home-page-tab-title-wrapper');
    if (!wrapper) {
      console.warn('[DMS] .home-page-tab-title-wrapper not found.');
      return;
    }

    // Remove the static placeholder tabs from the HTML template
    wrapper.innerHTML = '';

    libraries.forEach((lib, index) => {
      const tab = document.createElement('div');
      tab.setAttribute('data-tab-obj', toTabKey(lib.Title));
      tab.className = `home-page-tab-title flex-fill text-center${index === 0 ? ' home-page-tab-title-active' : ''}`;
      tab.textContent = lib.Title;
      wrapper.appendChild(tab);
    });
  }

  // ─────────────────────────────────────────────────────────────────────────
  // SharePoint REST
  // ─────────────────────────────────────────────────────────────────────────

  /**
   * Fetches all non-default document libraries and builds the full payload.
   *
   * Each library's ROOT nodes = its first-level folders.
   * The library name itself is never shown as a node — only its children are.
   */
  private async _getTabGridData(): Promise<{ payload: TabGridPayload; libraries: SPList[] }> {
    const libraries = await this._getDocumentLibraries();
    const payload: TabGridPayload = {};

    await Promise.all(
      libraries.map(async (lib) => {
        const key = toTabKey(lib.Title);

        // Resolve view ID for AllItems.aspx redirect URLs
        let viewId: string = lib.DefaultView?.Id || '';
        if (!viewId) {
          viewId = await this._getDefaultViewId(lib.Id);
        }

        // First-level folders of this library = root nodes in the UI column 1
        const firstLevelFolders = await this._getSubFolders(lib.RootFolder.ServerRelativeUrl);
        // console.log(`[DMS] "${lib.Title}" (${key}) → ${firstLevelFolders.length} root folder(s)`);

        const rootNodes: TreeNode[] = await Promise.all(
          firstLevelFolders.map(async (folder): Promise<TreeNode> => {
            const children = await this._getFolderChildren(
              folder.ServerRelativeUrl,
              viewId,
              1  // depth 1 = items directly inside this first-level folder
            );
            return {
              id: toId(folder.Name),
              label: folder.Name,
              iconType: 'FolderIcon',
              children,
            };
          })
        );

        payload[key] = rootNodes;
      })
    );

    return { payload, libraries };
  }

  private async _getDocumentLibraries(): Promise<SPList[]> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    const endpoint =
      `${siteUrl}/_api/web/lists` +
      `?$filter=BaseTemplate eq 101 and Hidden eq false` +
      `&$select=Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl,DefaultView/Id` +
      `&$expand=RootFolder,DefaultView`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(`Failed to fetch libraries: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const lists: SPList[] = data.value || [];

    return lists.filter((l) => EXCLUDED_LIBRARIES.indexOf(l.Title) === -1);
  }

  private async _getDefaultViewId(listId: string): Promise<string> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const endpoint = `${siteUrl}/_api/web/lists('${listId}')/DefaultView?$select=Id`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        return data.Id || '';
      }
    } catch (e) {
      console.warn(`[DMS] Could not fetch default view for list ${listId}`, e);
    }
    return '';
  }

  private async _getFolderChildren(
    folderServerRelativeUrl: string,
    defaultViewId: string,
    depth: number
  ): Promise<TreeNode[]> {

    const [subFolders, files] = await Promise.all([
      this._getSubFolders(folderServerRelativeUrl),
      this._getFiles(folderServerRelativeUrl),
    ]);

    const folderNodes: TreeNode[] = await Promise.all(
      subFolders.map(async (folder): Promise<TreeNode> => {

        if (depth >= MAX_DEPTH) {
          const hasSubFolders = await this._hasSubFolders(folder.ServerRelativeUrl);

          if (hasSubFolders) {
            // Max depth WITH sub-folders → redirect link to AllItems.aspx
            return {
              id: folder.UniqueId,
              label: folder.Name,
              type: 'link',
              iconType: 'FolderIcon',
              url: this._buildSharePointFolderUrl(folder.ServerRelativeUrl, defaultViewId),
              target: '_blank',
              children: [],
            };
          } else {
            // Max depth WITHOUT sub-folders → folder node with file leaves
            const leafFiles = await this._getFiles(folder.ServerRelativeUrl);
            const fileLeafNodes: TreeNode[] = leafFiles.map((file): TreeNode => ({
              id: file.UniqueId,
              label: file.Name,
              type: 'link',
              iconType: 'linkIcon',
              url: this._buildFileUrl(file),
              target: '_blank',
              children: [],
            }));
            return {
              id: folder.UniqueId,
              label: folder.Name,
              iconType: 'FolderIcon',
              children: fileLeafNodes,
            };
          }
        }

        // Normal case — recurse one level deeper
        const children = await this._getFolderChildren(
          folder.ServerRelativeUrl,
          defaultViewId,
          depth + 1
        );
        return {
          id: folder.UniqueId,
          label: folder.Name,
          iconType: 'FolderIcon',
          children,
        };
      })
    );

    const fileNodes: TreeNode[] = files.map((file): TreeNode => ({
      id: file.UniqueId,
      label: file.Name,
      type: 'link',
      iconType: 'linkIcon',
      url: this._buildFileUrl(file),
      target: '_blank',
      children: [],
    }));

    return [...folderNodes, ...fileNodes];
  }

  private async _hasSubFolders(folderServerRelativeUrl: string): Promise<boolean> {
    try {
      const sub = await this._getSubFolders(folderServerRelativeUrl);
      return sub.length > 0;
    } catch {
      return false;
    }
  }

  private async _getSubFolders(folderServerRelativeUrl: string): Promise<SPFolder[]> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const encodedPath = encodeURIComponent(folderServerRelativeUrl);

    const endpoint =
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedPath}')/Folders` +
      `?$select=UniqueId,Name,ServerRelativeUrl` +
      `&$filter=Name ne 'Forms'`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) {
        console.warn(`[DMS] Sub-folders failed for ${folderServerRelativeUrl}: ${response.status}`);
        return [];
      }
      const data = await response.json();
      return (data.value as SPFolder[]) || [];
    } catch (error) {
      console.warn(`[DMS] Error fetching sub-folders for ${folderServerRelativeUrl}:`, error);
      return [];
    }
  }

  private async _getFiles(folderServerRelativeUrl: string): Promise<SPFile[]> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const encodedPath = encodeURIComponent(folderServerRelativeUrl);

    const endpoint =
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedPath}')/Files` +
      `?$select=UniqueId,Name,ServerRelativeUrl,LinkingUrl`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) {
        console.warn(`[DMS] Files failed for ${folderServerRelativeUrl}: ${response.status}`);
        return [];
      }
      const data = await response.json();
      return (data.value as SPFile[]) || [];
    } catch (error) {
      console.warn(`[DMS] Error fetching files for ${folderServerRelativeUrl}:`, error);
      return [];
    }
  }

  // ── URL builders ───────────────────────────────────────────────────────────

  private _buildSharePointFolderUrl(folderServerRelativeUrl: string, defaultViewId: string): string {
    const origin = new URL(this.context.pageContext.web.absoluteUrl).origin;
    const siteRelativePath = this.context.pageContext.web.serverRelativeUrl; // e.g. /sites/DevDMS

    // Extract library name: /sites/DevDMS/Department Functions/Sales → "Department Functions"
    const withoutSite = folderServerRelativeUrl
      .replace(siteRelativePath, '')
      .replace(/^\//, '');
    const libraryName = withoutSite.split('/')[0] || '';

    const encodedLibraryName = encodeURIComponent(libraryName);
    const encodedFolderPath  = encodeURIComponent(folderServerRelativeUrl);
    const viewParam = defaultViewId ? `&viewid=${encodeURIComponent(defaultViewId)}` : '';

    return (
      `${origin}${siteRelativePath}/${encodedLibraryName}/Forms/AllItems.aspx` +
      `?id=${encodedFolderPath}${viewParam}`
    );
  }

  private _buildFileUrl(file: SPFile): string {
    if (file.LinkingUrl) return file.LinkingUrl;
    const origin = new URL(this.context.pageContext.web.absoluteUrl).origin;
    return `${origin}${file.ServerRelativeUrl}`;
  }

  // ── SPFx lifecycle ────────────────────────────────────────────────────────

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(() => { /* extend if needed */ });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
        let environmentMessage = '';
        switch (context.app.host.name) {
          case 'Office':
            environmentMessage = this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
            break;
          case 'Outlook':
            environmentMessage = this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
            break;
          case 'Teams':
          case 'TeamsModern':
            environmentMessage = this.context.isServedFromLocalhost
              ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
            break;
          default:
            environmentMessage = strings.UnknownEnvironment;
        }
        return environmentMessage;
      });
    }
    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}