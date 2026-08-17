# LaTeX Blocks homepage

The product homepage for LaTeX Blocks, built with vinext and the OpenAI Sites
runtime. Content lives in `app/page.tsx`; the visual system is in
`app/globals.css`.

## Local development

```bash
npm install
npm run dev
npm run build
```

The production build emits a Cloudflare Worker-compatible bundle under `dist/`.
