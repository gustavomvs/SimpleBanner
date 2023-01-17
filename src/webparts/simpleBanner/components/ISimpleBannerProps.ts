export interface ISimpleBannerProps {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  context: any;
  itemId: number;
  updatePropety: (id: number) => void;
  fileName: string;
  fileSize: number;
  updateFileName: (filename: string) => void;
  updateFileSize: (filesize: number) => void;
}
