---
description: Generate Sunday School material from a YouTube sermon
argument-hint: <youtube-url> <sermon-timespan> e.g. https://www.youtube.com/watch?v=XYZ 32:38-1:05:21
allowed-tools: Bash, Read, Write, Edit, Glob, Grep, WebFetch, WebSearch, Agent
audio download website: https://media.ytmp3.gg/
---

# Sunday School Material Generator

You are generating a middle school Sunday School lesson plan based on a YouTube sermon.

## Arguments

The user provides: $ARGUMENTS

Parse the arguments:
- **YouTube URL**: the YouTube video URL (e.g. `https://www.youtube.com/watch?v=...`)
- **Sermon time span**: start and end timestamps separated by a dash (e.g. `32:38-1:05:21`)

If arguments are missing, ask the user for them.

## Step 0: Check for User-Provided Materials

Before downloading anything, check if the user has provided additional source materials (transcript documents, PowerPoint files, etc.) in a specified folder (e.g., `C:\Users\jizjin\Downloads`). These often contain the speaker's intended content — correct scripture references, key phrases, sermon outline — and should be combined with the subtitle transcript for maximum accuracy.

Extract text from `.docx` files using `python-docx` and from `.pptx` files using `python-pptx` (Python 3.7 at `/c/Program Files/Python37/python`).

## Step 1: Obtain Subtitles

### Option A: Download from YouTube (preferred)

Use `yt-dlp` to download the auto-generated English subtitles from the YouTube video:

```bash
yt-dlp --write-auto-sub --sub-lang en --skip-download --sub-format srt -o "%(title)s" "<youtube-url>"
```

This will create an `.en.vtt` or `.en-orig.srt` file. Find the subtitle file that was created.

### Option B: Generate Locally via Audio + Whisper (fallback)

If YouTube auto-subtitles are not available (yt-dlp reports "no subtitles for the requested languages"), follow this fallback:

1. **Extract audio** from the YouTube video:
   ```bash
   yt-dlp -x --audio-format mp3 -o "%(title)s.%(ext)s" "<youtube-url>"
   ```
   If ffmpeg is not installed, install it: `winget install --id Gyan.FFmpeg -e`
   Then convert the downloaded webm to mp3: `ffmpeg -i input.webm -vn -ab 192k -ar 44100 output.mp3`

2. **Download the whisper model** (if not already present):
   ```bash
   curl -L -o ggml-base.bin "https://huggingface.co/ggerganov/whisper.cpp/resolve/main/ggml-base.bin"
   ```

3. **Generate subtitles** using ffmpeg's built-in whisper filter. **Important:** On Windows, run this via `cmd.exe` or a `.bat` file to avoid Git Bash path escaping issues:
   ```bat
   @echo off
   cd /d C:\alanjin14\sunday_school
   "C:\Users\jizjin\AppData\Local\Microsoft\WinGet\Links\ffmpeg.exe" -y -i "input.mp3" -af "whisper=model=ggml-base.bin:language=en:format=srt:destination=output.srt" -f null NUL
   ```

4. **Clean up** the whisper model file after transcription to save space.

## Step 2: Extract the Sermon Portion

Parse the subtitle file (SRT/VTT format) and extract only the lines that fall within the user-specified time span. If no time span is provided, use the full subtitle file (the sermon portion can be identified by context — worship music vs. spoken sermon).

To convert timestamps:
- `MM:SS` format means `00:MM:SS`
- `H:MM:SS` format is as-is
- Compare each subtitle's start time against the sermon start/end times

Read the subtitle file content and extract the text for the sermon time range. Clean up the text:
- Remove duplicate lines (auto-subs often repeat)
- Remove formatting tags like `<font>`, `[Music]`, etc.
- Consolidate into readable paragraphs

**Combine sources:** Use the subtitle transcript for actual spoken content (examples, illustrations, humor, audience interaction) AND the user-provided docx/pptx for intended content (correct scripture, key phrases, outline structure). The combination produces far more accurate materials than either source alone.

## Step 3: Generate Sermon Summary

Based on the extracted sermon transcript, create a sermon summary markdown file similar to the format in this project. Name it descriptively based on the sermon topic. Include:

- Source YouTube URL
- Church name: Evangelical Community Church (ECC) Redmond
- Date (derive from the video title or ask the user)
- Service structure (worship songs mentioned, prayer, sermon)
- Key Scripture passages
- Sermon summary with main points (typically 2-4 main points)
- Application section
- Any closing hymns mentioned

Save this as a separate summary file in the project root.

## Step 4: Generate Sunday School Lesson Plan

Using the sermon summary AND the extracted transcript, generate a complete Sunday School lesson plan for **middle schoolers (Grades 6-8)** in a **1-hour session**.

Use this EXACT schedule structure:

| Time | Activity | Duration |
|------|----------|----------|
| 11:15 | Icebreaker Game | 8 min |
| 11:23 | Sermon Review & Teaching (with short video clip) | 10-12 min |
| 11:35 | Transition to Small Groups | 2-3 min |
| 11:38 | **Small Group Discussion** | **30 min** |
| 12:18 | Wrap-Up & Prayer | 2 min |

The lesson plan MUST include all of these sections:

### LEADER'S OVERVIEW
- Main Scripture reference
- Theme (one sentence)
- Goal for students

### ICEBREAKER GAME (8 min)
- A creative, themed game that connects to the sermon topic
- Clear instructions for how to play
- Transition question that bridges to the sermon

### SERMON REVIEW & TEACHING (10-12 min)
- Read the key scripture together
- 2-3 Key Ideas from the sermon, simplified for middle schoolers
- Each key idea should be concise (3-5 sentences) and relatable
- Optional short video suggestion if relevant

### SMALL GROUP DISCUSSION (40 min)
- Split into groups of 4-5 with one leader
- Include these question categories:
  1. **Warm-Up Questions** (2-3 questions) - what students remember from the sermon
  2. **Understanding the Topic** (2-3 questions) - dig into the main ideas
  3. **Faith vs Real Life** (2-3 questions) - connect theology to daily experience
  4. **Real Life Application** (2 questions) - practical middle school scenarios
  5. **Personal Reflection** (2 questions) - introspective, what will they do differently

### WRAP-UP & PRAYER (2 min)
- A sentence completion exercise
- A closing prayer related to the sermon theme

## Style Guidelines

- Use language appropriate for middle schoolers (Grades 6-8)
- Include emoji sparingly in lists (like the icebreaker) to keep it engaging
- Use bold text for emphasis on key theological terms
- Include scripture quotes in blockquotes
- Use horizontal rules (---) between major sections
- The tone should be warm, engaging, and not preachy
- Connect theological concepts to real middle school experiences (school pressure, friendships, social media, identity)

## Step 5: Generate Presentation Slides

Create a PowerPoint-style HTML slide deck for the **Icebreaker Game** and **Sermon Review & Teaching** sections. Save it as `ECC Sunday School - <Sermon Title> (Slides).html` in the project root.

The slide deck should be a self-contained HTML file with embedded CSS and JavaScript for navigation (arrow keys and click to advance).

### Slide Design Requirements

- **Full-screen slides** with a clean, modern look suitable for projecting in a classroom
- **Large readable text** (minimum 32px for body text, 48px+ for headings)
- **Dark background with light text** (e.g., dark navy/charcoal with white text) for easy reading on projectors
- Use a consistent color scheme throughout
- Each slide should have minimal text — bullet points, not paragraphs

### Icebreaker Slides

- **Title slide**: Game name and brief instructions
- **One slide per icebreaker item/scenario** — large text, with a relevant emoji
- **Transition slide**: The bridge question that connects the game to the sermon topic

### Sermon Review & Teaching Slides

- **Scripture slide**: Display the key verse(s) in large text with the reference
- **One slide per Key Idea** — title + 2-3 short bullet points max
- Include relevant images where helpful. To include images, use one of these approaches:
  - **YouTube video thumbnails**: For the sermon video or related Bible Project videos, embed the thumbnail using `https://img.youtube.com/vi/<VIDEO_ID>/maxresdefault.jpg`
  - **Search for relevant free images**: Use WebSearch to find appropriate Creative Commons or public domain images (e.g., from Unsplash, Pixabay) that illustrate key concepts. Embed them via their direct URL.
- **Video slide**: If a short video clip is recommended (e.g., Bible Project), include a slide with:
  - A clickable YouTube link or embedded thumbnail that links to the video
  - Brief description of what to watch and how long

### Slide Navigation

- Arrow keys (left/right) to navigate between slides
- Click anywhere to advance
- Show a small slide counter (e.g., "3 / 12") in the bottom-right corner
- Support fullscreen mode (F key to toggle)

## Output

Save three files in the project root:
1. **Summary file**: `ECC Redmond English Sunday Service <date> - Summary.md`
2. **Lesson plan file**: `ECC Sunday School - <Sermon Title> (Middle School).md`
3. **Slides file**: `ECC Sunday School - <Sermon Title> (Slides).html`

After saving, show the user a brief summary of what was generated and the file names.
