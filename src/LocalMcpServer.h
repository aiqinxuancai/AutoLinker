#pragma once

namespace LocalMcpServer {

void Initialize();
void Shutdown();
bool IsRunning();
int GetBoundPort();

} // namespace LocalMcpServer
