# gfm-validate
A simple GitHub Flavored Markdown (Gfm) validator. Uses the [Octokit .NET client library for GitHub](https://github.com/octokit/octokit.net). 

![screenshot](https://github.com/AndrewJByrne/gfm-validate/blob/master/assets/app-screenshot.png)

This Windows desktop app demonstrates the following:
* Connecting to GitHub APIs using the Octokit library for .NET 4.5
* Parsing markdown into raw HTML using the RenderRawMarkdown method from the Octokit library
* Displayig HTML using the GitHub CSS
* A little bit of MVVM
* Handling drag/drop of markdwon files onto the app surface
* App settings
* Suppression of browser control warnings, preventing browser control from being a drop target

## Copyright and License

Copyright 2015 Andrew J. Byrne

Licensed under the [MIT License](https://github.com/AndrewJByrne/gfm-validate/blob/master/LICENSE)
