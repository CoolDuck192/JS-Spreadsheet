import { Spreadsheet } from "../App";

export function SpreadsheetEmbeddingDemo() {
  return (
    <div data-testid="embedding-host" style={{ height: 520 }}>
      <Spreadsheet storage={false} style={{ height: 480 }} />
    </div>
  );
}
