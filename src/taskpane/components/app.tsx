import type { FC } from "react";
import { ChatInterface } from "./chat";
import { ErrorBoundary } from "./error-boundary";

interface AppProps {
  title: string;
}

const App: FC<AppProps> = () => (
  <ErrorBoundary>
    <div className="h-screen w-full overflow-hidden">
      <ChatInterface />
    </div>
  </ErrorBoundary>
);

export default App;
