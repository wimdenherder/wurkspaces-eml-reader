export interface FileSelectedEvent {
  userTimezone: UserTimezone;
  commonEventObject: CommonEventObject;
  drive: Drive;
  hostApp: string;
  clientPlatform: string;
  userLocale: string;
  userCountry: string;
}

interface UserTimezone {
  id: string;
  offSet: string;
}
interface CommonEventObject {
  userLocale: string;
  platform: string;
  hostApp: string;
  timeZone: Timezone;
}

interface Timezone {
  id: string;
  offset: number;
}
interface Drive {
  selectedItems: SelectedItem[];
  activeCursorItem: SelectedItem;
}

interface SelectedItem {
  title: string;
  id: string;
  mimeType: string;
  iconUrl: string;
}