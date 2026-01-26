/**
 * OneDrive Files Types
 */

export interface DriveItem {
  id: string;
  name: string;
  size: number;
  webUrl: string;
  lastModifiedDateTime: string;
  createdDateTime: string;
  folder?: { childCount: number };
  file?: { mimeType: string };
  parentReference?: {
    id: string;
    path: string;
    driveId: string;
  };
  '@microsoft.graph.downloadUrl'?: string;
}

export interface DriveItemSummary {
  id: string;
  name: string;
  size: number;
  isFolder: boolean;
  mimeType?: string;
  webUrl: string;
  modified: string;
}
