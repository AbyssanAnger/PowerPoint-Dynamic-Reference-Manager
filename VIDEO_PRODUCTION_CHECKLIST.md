# Video Production Checklist

## 1. Pre-Production (Vorbereitung)
- [ ] **Build Add-In**: Ensure `ReferenceManager.ppam` is up-to-date (Follow `src/BUILD_ADDIN.md`).
- [ ] **Demo Assets**:
    - [ ] A "Messy" PowerPoint presentation (slide with manual numbers like `[1]`, `[3]` out of order).
    - [ ] A "Clean" PowerPoint presentation (slides with `[Key]` placeholders).
    - [ ] A dummy `.bib` file (e.g., `test_references.bib`).
- [ ] **Clean Desktop**: Hide desktop icons or record only the PowerPoint window to look professional.

## 2. Screen Recording (OBS or similar)
Record these clips separately to make editing easier:
- [ ] **Clip A (The Struggle)**: You manually deleting a number, typing a new one, scrolling back up to check a number. (Action: Frustrated cursor movement).
- [ ] **Clip B (The Solution)**: You clicking the "Update References" button. Citations change instantly. Bibliography appears.
- [ ] **Clip C (Installation)**: 
    - Open PPT -> File -> Options -> Add-Ins -> Go -> Add New -> Select `.ppam`.
    - *Note:* Do this slowly and smoothly; we will speed it up in post.
- [ ] **Clip D (GitHub & Outro)**: Scroll through your GitHub repo, show the "Code" button.

## 3. Audio Recording
- [ ] Record the voiceover using the script (`video_script.md`) in one take or paragraph-by-paragraph.
- [ ] Use a decent microphone. If not available, use your phone's memo app in a quiet room (closets works great for dampening echo!).

## 4. Editing in DaVinci Resolve 20
- [ ] **Import**: Drag all video clips and your audio file into the Media Pool.
- [ ] **Audio Base**: Drag the voiceover to the timeline (Audio Track 1) first. Cut out silences/breaths.
- [ ] **Video Sync**: Drag video clips to Video Track 1 to match what you are saying.
- [ ] **Speed Ramping (Installation)**: 
    - Right-click the Installation clip -> `Retime Controls` (Ctrl+R).
    - Drag the speed arrow to 200% or 400% so it fits the short "It's super easy to install" sentence.
- [ ] **Zoom/Pan**: Use "Dynamic Zoom" or Keyframes in the "Inspector" to zoom in on small buttons (like the Ribbon button or specific citations).
- [ ] **Background Music (Optional)**: Add low-volume lo-fi or corporate background music (Audio Track 2, set volume to -20dB or lower).

## 5. Export
- [ ] Go to "Deliver" tab.
- [ ] Preset: YouTube 1080p or Custom (H.264/H.265, MP4).
- [ ] Render!
