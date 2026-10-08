# Photo Montage

Turns a batch of photos into a finished video for a conference screen. Slow Ken Burns motion, crossfades, exact timing and a loop that joins without a jump. Everything runs in the browser. No photo is uploaded.

Lives at `/photo-montage.html`. The homepage links to it under Other Tools and passes on the pixel size from the calculator if one has been entered. It used to be called Slideshow Builder, and `/slideshow.html` now forwards to the new address so old links keep working.

## Using it

1. Drop photos on the page or press Add photos.
2. Pick a screen size or type the width and height.
3. Set the time per photo, or switch to Total duration and give the length of the whole video.
4. Leave Seamless loop on for a file that repeats, or turn it off for one that plays once.
5. Press play to preview. Check loop point plays the last few seconds straight into the start.
6. Press Export video, then Download video.

The defaults are 1920 x 1080 at 30 fps, 5 seconds a photo with a 1 second crossfade, gentle motion, looping on, unsuitable photos excluded and H.264 in MP4 at High quality where the browser can encode it.

## How timing works

All timing is in whole frames.

A photo's time runs from the frame it is fully on screen to the frame the next photo is fully on screen. The crossfade into the next photo is the last part of that time. It is never added on top. So the video length is the photo times added up, and 20 photos at 5 seconds make exactly 1:40.00 whatever the crossfade length.

A crossfade can use at most half of a photo's time. If it is set longer it is shortened and the page says so. A crossfade of 0 seconds is a cut.

In Total duration mode the total is shared equally. If it does not divide evenly some photos get one frame more than others, spread through the run. A photo with its own duration keeps it and the rest share what is left.

If a duration is not a whole number of frames at the chosen frame rate it is rounded to the nearest frame and the page shows the exact figure used.

## How the loop works

The timeline is periodic. With N frames in the video, frame N would be frame 0 again, so it is never rendered. N frames go into the file and no copy of the first frame is added at the end.

The last photo's crossfade leads into the first photo like every other crossfade and is counted in the length. The file ends as that crossfade completes and starts with the first photo fully on screen. The first photo's motion begins during the closing crossfade and carries on across the join, so the last frame and the first frame are neighbours on one continuous move.

With one photo the motion goes out and comes back with a smooth turn at each end, so there is nothing to reset.

With cuts the loop point is a cut from the last photo to the first, the same as any other cut in the sequence. The page says so.

Check the loop in a player set to repeat. Some players pause for a moment while they reopen a file. That pause comes from the player, not the video.

## Which photos are used

Exclude photos that require excessive cropping is on by default.

Each photo is cropped to fill the screen, then tightened by the most its motion will zoom. If less than 55% of the photo is still on screen at that point it is left out. The figure is under Photo fit, Advanced. Portrait photos drop out of a landscape video and landscape photos out of a portrait one. A 3:2 or 4:3 photo on a 16:9 screen stays.

Excluded photos stay in the project and are listed with the reason. Any of them can be put back, cropped to fill or shown whole over a blurred background. A choice made by hand is kept when the screen size changes.

With the option off every photo is used and there are three ways to show them

- Fill crops to cover the screen
- Fit shows the whole photo on a background colour
- Blurred background shows the whole photo over a blurred and darkened copy of itself

Photos are never stretched.

## Motion

Pushes zoom in around the focus point, pulls are the reverse and pans slide along whichever side of the photo already overhangs the screen. How far a photo travels is capped per second as well as in total, so a short photo moves less and not faster.

Intensity sets the most a photo zooms over its time on screen. Subtle is 4%, Gentle is 7% and More is 12%.

Vary the moves automatically changes the order of moves, pan directions and a slight drift. The result is fixed until New variation is pressed. Preview and export always use the same moves.

Select a photo and open Advanced settings for this photo to choose its move, set a focus point, set its own start and end framing by eye, give it its own duration or change the transition after it. Reset this photo to automatic clears all of that.

## Order

Photos import sorted by file name. Drag to reorder, or select a photo and use Earlier and Later, or hold Alt and press the arrow keys. Randomise order shuffles with a new seed and then stores the order, so replaying or exporting never reshuffles. Restore import order puts it back.

## Formats

The page asks the browser what it can encode at the exact size, frame rate and bitrate chosen. Formats it cannot encode are listed but cannot be picked. If the selected format stops being available after a change of size, export is blocked with the reason. Nothing is substituted or resized without being asked.

| Format | Where it plays | Notes |
| --- | --- | --- |
| H.264 in MP4 | PowerPoint, Keynote, QLab, VLC, most show playback | The default. High profile, 8 bit 4:2:0. Needs an even width and height |
| H.265 in MP4 | Recent Macs and some Windows machines | Only offered where the browser can encode it. Test on the playback machine |
| Apple ProRes in MOV | QLab, Resolume, Millumin, disguise, Watchout, editing software, Macs | Proxy, LT, 422 or HQ. 10 bit 4:2:2. Not offered by default |
| VP9 in WebM | Browsers and VLC | Accepts odd sizes |
| AV1 in MP4 or WebM | Recent players only | Accepts odd sizes |

### Bitrate

For everything except ProRes, the bitrate slider sets the rate the encoder aims at, from 1 to 300 Mbit/s. The presets set it for you and scale with the picture size and frame rate. At 1920 x 1080 and 30 fps H.264 gets 12, 25 or 50 Mbit/s for Standard, High and Maximum. Moving the slider by hand fixes the rate until a preset is pressed again.

Crossfades are where a low bitrate shows first, because every pixel changes on every frame. Measured on a Mac's hardware H.264 encoder with detailed, grainy test images, picture quality rose steadily from 8 to 80 Mbit/s with no point where it levelled off, so if fades look blocky, raise the rate. Above 60 Mbit/s the VT Inspector flags H.264 as heavy for PowerPoint on Windows. Media servers do not mind.

Bitrate mode, key frame spacing and encoder preference are under Advanced. The estimated file size is bitrate times duration. Montages often come in smaller because the encoder only uses what it needs.

### ProRes

ProRes is not something browsers can encode, so it is done with prores-wasm-encoder, a WebAssembly build of a ProRes encoder, spread across up to 8 of the computer's cores. It takes noticeably longer than H.264. Every frame stands alone, so there is no bitrate to set and nothing to smear in a crossfade. Pick the profile instead. At 1920 x 1080 and 29.97 fps they run at about 45, 102, 147 and 220 Mbit/s for Proxy, LT, 422 and HQ, so an HQ minute is about 1.65 GB. Files that size are written straight to disk in Chrome and Edge.

HAP, DNxHR and NotchLC cannot be made in a browser. If a media server needs one of them, export ProRes and transcode that.

Colour is converted on the page with the BT.709 matrix for anything 1280 wide or more or over 576 lines, and BT.601 below that. The file is tagged to match. Left alone, Chrome uses BT.601 for every size, which the VT Inspector flags on an HD picture.

ProRes frames are converted and tagged BT.709 by the encoder library at every size.

After every export the page reads the finished file back and shows the frame count, picture size, frame rate, duration and colour it actually contains next to what was asked for. MP4 files also go through the VT Inspector engine for a PowerPoint on Windows verdict. Chrome cannot play ProRes, so a ProRes file's playback check says it was not checked.

## Browser requirements

A desktop browser with WebCodecs, or for ProRes only, WebAssembly and module workers. Chrome or Edge is the tested route. Safari 16.4 and Firefox 130 or later have the same interfaces but have not been tested with this page. The page must be served over https or from localhost. Phones are not supported.

Which codecs appear depends on the browser and the machine. Google Chrome and Edge include an H.264 encoder. Open source Chromium builds do not.

## Limits worth knowing

- Frame rates are 24, 25, 30, 50 and 60. There are no fractional rates such as 29.97.
- The canvas stops at 16384 px a side. Beyond that the encoder decides, and H.264 hardware on many machines stops near 4096 x 2304.
- The finished video is built in memory unless it is written straight to disk. In Chrome and Edge large exports ask where to save first and write to disk as they go. That starts at an estimated 1 GB and can be forced either way under Advanced.
- WebM stores times in whole milliseconds, so frame times in a WebM file are rounded to the millisecond. The frame count is still exact.
- Transparent areas in a PNG or WebP show black in Fill, the background colour in Fit and the blur in Blurred background.
- No audio, captions, titles or logos.
- Warnings about heavy settings are estimates of how demanding the job is. The browser does not say how much memory it will allow, so they cannot predict a failure.

## Files

| File | What it is |
| --- | --- |
| `photo-montage.html` | The page |
| `slideshow.html` | Forwards the old address to the new one |
| `slideshow-app.js` | Page logic and interface. The script files kept their old names |
| `slideshow-core.js` | Timeline, motion, suitability, ordering and codec maths. No DOM, runs in Node |
| `slideshow-render.js` | Draws one frame. Used by preview and export alike |
| `slideshow-export.js` | Capability checks, colour conversion, encode loop |
| `slideshow-worker.js` | Runs the export off the page |
| `mp4-muxer.js`, `webm-muxer.js` | Container writers by Vanilagy, MIT licence, vendored. `webm-muxer.js` carries one marked change so the segment duration includes the last frame |
| `prores-encoder-parallel.min.js` | ProRes encoder by Olivier Estevez, LGPL 2.1, vendored unmodified. Licence in `prores-encoder.LICENSE.txt` |
| `media-inspect.js` | Existing VT Inspector engine, reused to check MP4 and MOV output |
| `tests/` | Not needed on the site |

Adding `?export=page` to the address runs the export on the page instead of in a worker. The page does this by itself if a worker cannot start. Adding `?w=3840&h=1080&fps=50` presets the output.

## Tests

From the repository root

```
node --test
```

64 tests with no dependencies. They cover timing in both modes, rounding, overrides, transition limits, the exclusion rule in both orientations, reversibility, seeded ordering, loop periodicity and continuity across the loop point for one, two and many photos, clean non looping ends, validity of every position along every move, renderer output at the loop point, preview and export parity, codec levels and the colour conversion.

`tests/browser/` holds the checks that need a real browser. They export real files and measure them with ffmpeg. See the note at the top of `tests/browser/run.js`.
