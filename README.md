# resume-parser

A resume parser for two-column PDFs.

A lot of parsers read a page from left to right. On a two-column resume that weaves the sidebar into the jobs, so a skill list and a role end up on the same line. This one looks for a blank vertical gutter. When it finds one, it reads the left column from top to bottom, then the right column.

A single-column PDF has no gutter, so the page is read straight down. That layout is the easy case. The gutter split is why this exists.

DOCX files follow the order stored in the document. The column split is for PDFs, where every character has a position.

## Install

Python 3.12 and [uv](https://docs.astral.sh/uv/).

```bash
git clone https://github.com/Logan1x/resume-parser.git
cd resume-parser
uv sync
uv run resume-parser resume.pdf
```

The command prints JSON.

## What comes out

Profile, work history, education, skills, projects, talks, and accomplishments.

`brief.experience` is how long the jobs add up to. Overlapping dates count once. `brief.gaps` lists stretches with no job. An empty list means the dated roles run together. A job that ends in December and another that starts in January counts as continuous. If the last role already ended, the months since then are a gap.

```json
{
  "brief": {
    "experience": "5 years 2 months",
    "gaps": [],
    "role": "Software Engineer · Northwind",
    "education": ["B.Tech, Computer Science · Example Institute"],
    "tech": ["Python", "PostgreSQL"]
  },
  "work": [
    {
      "company": "Northwind",
      "title": "Software Engineer",
      "date": "Jan 2021 – Present",
      "descriptions": ["Built the billing API."]
    }
  ]
}
```
