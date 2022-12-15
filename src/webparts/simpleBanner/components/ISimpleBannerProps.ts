export interface ISimpleBannerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  context: any;
  itemId: number;
  updatePropety: (id: number) => void;
}
