# Quiz Control Google Forms Add-on

This project provides the source for a Google Apps Script add-on that augments Google Forms quizzes with session control features. The add-on adds a configurable timer and enforces retakes for submissions that do not meet a minimum passing score.

## Features

- **Configurable timer** – open the "Quiz Control" menu inside the form editor to choose the session duration and start a timed window. The add-on automatically closes the form when time expires and can cancel a running timer at any moment.
- **Forced retake** – enable the retake option to automatically delete submissions below the configured passing threshold (80% by default) and send an email prompting the respondent to try again.
- **Sidebar configuration** – update timer duration, passing threshold, and retake settings from the add-on sidebar.

## Project structure

| File | Description |
| ---- | ----------- |
| `Code.gs` | Core Apps Script logic for the add-on, including menu creation, timer automation, and quiz retake enforcement. |
| `TimerConfig.html` | Sidebar user interface that lets editors configure timer and retake preferences. |
| `appsscript.json` | Add-on manifest declaring triggers and metadata for a Forms add-on. |

## Deployment

1. Visit [script.google.com](https://script.google.com) and create a **new standalone Apps Script project**.
2. Replace the default files with the ones in this repository. Enable the "Show appsscript.json" setting so the manifest can be updated.
3. Deploy the project as an **add-on test deployment** (Deploy → Test deployments → Add-on). Install the test add-on in your account.
4. Open any Google Form, then use **Add-ons → Quiz Control → Configure settings** to choose a timer duration and passing score for that form.
5. Start a timed session from the add-on menu. The add-on stores settings per form, so you can reuse the same deployment across multiple Forms files.

> **Note:** Forced retake requires the quiz to collect respondent email addresses and use auto-graded questions so that a total score is available immediately on submission.
