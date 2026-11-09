/** The full-page app: the editor with the options and page connections of the shell. */
import { SumlineEditor } from '@sumline/react';
import { useEffect, useState } from 'react';
import { connectShell, shellOptions, watchLaunches } from '@sumline/web-shared/shell';

export function App() {
  const [options, setOptions] = useState(shellOptions);
  // A file opened through the installed app replaces the workbook.
  useEffect(() => watchLaunches((handle) => setOptions((o) => ({ ...o, document: handle }))), []);
  return <SumlineEditor {...options} onReady={connectShell} />;
}
