# Third-party components

wordplai uses Windows/.NET Framework and the installed Microsoft Word
COM object model. Neither Microsoft Word nor a separate Office distribution is
included. The add-in sources use only framework libraries; no NuGet package is
required at runtime.

The installer includes [zrok2 2.0.4](https://github.com/openziti/zrok/releases/tag/v2.0.4)
(upstream build 6ff92039), under the Apache License, Version 2.0. Its
license is included as `zrok-LICENSE`. Payload hashes are recorded in
`build-manifest.json`. zrok/OpenZiti transitive dependencies retain their
respective upstream licenses; the upstream module metadata is available in the
source release and through `go version -m runtime/zrok.exe`.

The pinned official Windows archive, extracted executable and license SHA-256
values are in `packaging/zrok.json`. `packaging/fetch-zrok.ps1` verifies those
inputs; the installer contains no upstream account profile or enrollment state.

This project was extracted from the Word module in UnderTree. Legacy identifiers
remain for upgrade compatibility; the other application is not a runtime or
build dependency. The source package declares `UNLICENSED`; these third-party
notices do not grant a license to the project's own code.

The embedded ChatGPT document picker bundles @chenglou/pretext 0.0.8 from the
repository lockfile, under the MIT License:

MIT License

Copyright (c) 2026 Pretext contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
