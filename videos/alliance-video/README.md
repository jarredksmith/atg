# Alliance Technical Group — LinkedIn Marketing Video

A HyperFrames video composition. Plain HTML + GSAP; rendered to MP4 by the `hyperframes` CLI.

25 seconds · 1920x1080 · 8 scenes · 3 shader transitions

## Requirements

- **Node.js 22+** — [nodejs.org](https://nodejs.org/)
- **FFmpeg** — `brew install ffmpeg` (macOS) or `sudo apt install ffmpeg` (Debian/Ubuntu) or [ffmpeg.org/download](https://ffmpeg.org/download.html) (Windows)

Verify: `npx hyperframes doctor`

## Preview

```bash
npx hyperframes preview
```

Opens the HyperFrames Studio at `http://localhost:3002` with frame-accurate scrubbing.

## Refine with Claude Code

This project was drafted in Claude Design. To polish animations, timing, and pacing:

```bash
npx skills add heygen-com/hyperframes   # install HyperFrames skills (one-time)
npx hyperframes lint                     # verify structure (should pass with zero errors)
npx hyperframes preview                  # open the studio for live feedback
```

Then open in Claude Code and iterate:

- "Make scene 5's hero entrance snappier"
- "Tighten the pacing — scene 4 feels too long"
- "Change the shader on transition 3 to light-leak"

## Render

Run from the project directory (the folder containing `index.html`).

```bash
npx hyperframes render -o output.mp4
```

1920x1080 / 30fps by default. Use `--fps 60` or `--resolution 3840x2160` to override.
