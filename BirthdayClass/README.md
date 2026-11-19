# BirthdayClass

A small, static birthday gift site made with HTML, CSS, and vanilla JavaScript. No build step or dependencies are required.

## Preview the site

Open `index.html` in a browser. The page uses ordinary deferred scripts and local files, so it can be previewed directly without a server.

## Change the recipient's name

Open `classes.js` and edit the `recipientName` value near the top. The name appears on the homepage and class pages.

## Add an instructor, portrait, and video

1. In `classes.js`, add an object to the `classes` array. Give it a unique `id`, and fill in the instructor, role, title, biography, description, lessons, portrait, and `videoUrl` fields. The existing six entries are examples to copy.
2. For a CSS portrait, set `portrait` to a new short name, then add a matching `.portrait--your-name` rule in `styles.css`. The portrait is assembled from CSS shapes, so no image files are needed. To use a real portrait instead, place the image in this folder and update the portrait markup/styles in `app.js` and `styles.css` to reference its relative path.
3. Set `videoUrl` to a relative video path, such as `media/lesson.mp4`. Add that file inside this project folder. Leave `videoUrl` empty to show the "Class coming soon" panel. Videos use native controls and do not autoplay.

Class pages use hashes such as `#class/pasta`, so browser Back and Forward work without server routing.
