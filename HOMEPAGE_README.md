# Weiyan Xie Academic Homepage

Personal academic homepage built with the [luost26/academic-homepage](https://github.com/luost26/academic-homepage) Jekyll template.

## Local preview

```bash
bundle install
bundle exec jekyll serve
```

Open `http://127.0.0.1:4000`.

## Deploy to GitHub Pages

1. Create a repository named `vaynexie.github.io`.
2. Push this folder to the repository.
3. Enable GitHub Pages (Settings → Pages → deploy from `main` branch).
4. Site URL: `https://vaynexie.github.io`

## Update content

- Profile, teaching, service, awards: [`_data/profile.yml`](_data/profile.yml)
- Research themes: [`_data/research_areas.yml`](_data/research_areas.yml)
- Publications source: [`cv_content.json`](cv_content.json)

Regenerate publication pages:

```bash
python3 generate_publications.py
```

Regenerate CV files:

```bash
./.venv/bin/python generate_cv_docx.py
cp research_cv_updated.pdf assets/cv/
cp research_cv_updated.docx assets/cv/
```

## Links

- Google Scholar: https://scholar.google.com/citations?user=_kc1e7EAAAAJ&hl=en
- GitHub: https://github.com/vaynexie
- Email: wxieai@cse.ust.hk
