# chess-feed

*Српски: see [README.sr.md](README.sr.md)*

An RSS feed of chess-results.com tournaments for a fixed list of FIDE IDs. This repository holds the **code and the
workflow** and publishes the feeds with GitHub Pages. The player list, the settings and the script's memory live in a
separate **private** repository.

## What it does

The script searches chess-results.com for each FIDE ID, downloads the Excel table that the site's player search
offers, and keeps the tournaments inside a time window: finished in the last `finished_days` days (default 33),
ongoing, and upcoming ones that finish within `upcoming_days` days (default 366).

When a player shows up in a tournament for the first time, an item is added to the RSS feed, for example
*"Играч Б је пријављен за: Турнир X"* (the verb follows `M`/`F`). A tournament discovered only after it finished reads
*"… је учествовао/учествовала на: …"*. If a player disappears from the starting list of an unfinished tournament and
stays missing on `withdraw_after_days` different days, one item *"… више није на стартној листи: …"* is added.
Nothing is announced twice. The text is Serbian Cyrillic and the links point to the Serbian (`lan=14`) chess-results page.

## Results of finished tournaments (optional)

Switched off by default (`[results] enabled = no` in `config.ini`). When on, the script waits until a tournament ended
more than `wait_days` days ago, shows all its rounds as played, and has not been changed on chess-results for
`stable_hours` hours. Then it opens the player's own page in that tournament (one request, at most `max_per_run` per
run). A feed item is added only if the page is really about that player (the FIDE ID is on it), every game has a
result, and at least `min_games` games were really played (byes and forfeits do not count), for example
*"Играч Ф је играо на: Турнир X – 5,5/9, 17. место"*, linking to the player's results
(`https://chess-results.com/tnrNNN.aspx?lan=14&art=9&snr=NN`). Nothing is announced if the player did not play, or if
the page cannot be verified.

- A tournament that is over but whose results are not final yet is tried again after `retry_hours` (default 4, one try
  per run). A page that cannot be read is tried again after `retry_other_hours` (default 24).
- When one page shows that the tournament was changed recently, the other players of the same tournament wait too, so a
  tournament with many players costs one request per retry, not one per player. Pages never tried before go first.
- After `give_up_days` days from the end date the tournament is given up (nothing is announced).
- For the earliest announcement set `wait_days = 0` and a small `stable_hours` (for example 3): the item then appears a
  few hours after the organizer uploads the final results. A correction made later cannot be announced.
- Tournaments found only after they ended are announced this way too, instead of as "је учествовао".
- A sign-up item has a body with the tournament's data too (same fields as below, without the player's data). One
  page is read per tournament and shared by all players, at most `[details] max_per_run` per run. If the page cannot be
  had, the announcement waits up to `[details] wait_hours` and is then sent without the data. The average rating and
  age is left out of sign-up items, because it changes until the start (`join_hide_average`); result items show the
  final value. Tournament data older than `cache_days` is read again.
- A result item has a body with the tournament's data (time control, location, number of rounds, type, date, average
  rating and age) and the player's data (starting rank, rating, rating performance, FIDE rating change, points,
  number of rounds played, place). All of it comes from the same page, so it costs no extra request.
- **Second check:** `recheck_days` (default 3) days after the tournament ended, and at least a day after the
  announcement, the page is read once more. If the player's results or stats changed, an item
  *"ИСПРАВКА: Играч је играо на: …"* is added, listing what changed. If nothing changed, nothing is added.
- **If chess-results does not answer** while a page is being read, that page is simply tried again after
  `retry_hours`; after two failures in a row the rest of the pages wait for the next run. A failure for one player
  never stops the others, and everything collected is kept and published.
- The saved state is written after every player and every page, so a stopped or crashed run keeps what it had found.
- Try it first with the workflow mode `check-results`, which shows what is read from up to three finished
  tournaments so you can compare it with the page in a browser.

## Active and inactive players

In `players.txt` (private repo) each line is `FIDE_ID  [A, I or R]  [M or F]  [groups]  label`, for example
`9213877 A M 1,3 Играч 3`. The flags may come in any order; `A` (active) and `M` are the defaults; `R` means retired. Groups are numbers
such as `1` or `1,3` (see below); a player without a group number joins the `default` group of `config.ini`.

- **Active** players are checked inside the active window (default 05:00 to 01:00, Central European time), at least
  4 hours apart: five checks a day.
- **Inactive** players are checked once a day inside the night window (default 01:00 to 05:00).
- **Retired** players (`R`) are checked once a month (`retired_day`, the 1st), in the same night window, after the
  inactive ones. If that night is missed they are checked the first night after it.
- New players are checked at the first run, whatever the time.

The workflow starts at six fixed UTC times a day (`cron:` in `watch.yml`); the script decides who is due, so most runs
are short. All numbers are in `config.ini` under `[schedule]`. Winter/summer time is handled by the script.

## How it is set up

| Repository | Visibility | Holds |
|---|---|---|
| this one (`chess-feed`) | public | `chess_watch.py`, `watch.yml`, the READMEs; the feeds are published from here with Pages |
| `chess-watch` | **private** | `config.ini`, `players.txt`, `state/` (what was announced, and the request log of the last 48 hours) |

Each run checks out the private repository next to the code, runs the script, publishes the feeds, and saves the state
back to the private repository. Standard GitHub-hosted runners are free for public repositories, so the number of runs
is not limited by a monthly allowance.

## Groups and feeds

Every group is one feed. The groups are defined in `config.ini` under `[groups]` as `number = file name | title`,
for example `1 = kikinda | ШК Раднички Кикинда`. A player can belong to several groups. To add a group, add a line.

- group feeds: `https://YOURNAME.github.io/chess-feed/kikinda.xml`, `srb.xml`, `fide.xml`
- one feed per player: `https://YOURNAME.github.io/chess-feed/fide-FIDEID.xml` (title from `[feed] player_title`)

After changing groups or titles, run the workflow in mode `rebuild`: it regenerates the feeds from the saved state
without checking any player. Feed files that no longer exist disappear with the next publication.

## Privacy and search engines

- The player list, settings and state are in the private repo. The public repo contains no player data, and the feeds
  are not committed to it (they are deployed straight to Pages).
- The run logs of this public repo show only player numbers (`player #3`), never IDs or names. The detailed log, with
  IDs and labels (every request, and one result line per player), is written to the private repo in
  `state/request_log.txt` and keeps the last 48 hours.
- The feeds do contain FIDE IDs and labels, by design. Nothing links to them, but anyone who knows or guesses the
  address can read them, and an XML file cannot carry a `noindex` tag, so a search engine could still pick it up. If
  that ever happens, a removal request can be filed with the search engine.

## Being gentle with chess-results.com

- One search form load per run, then one request per due player, with `delay_seconds` plus a random `jitter_seconds`
  between players. `max_players_per_run` caps a run.
- A `403` or `429` answer stops the whole run at once; three failures in a row also stop it.
- Failed checks never count as "player left the tournament".
- A search that finds nothing makes the site return a page instead of an Excel file; that is treated as "no rows", not
  as a failure. For a player with open tournaments, an empty answer is first confirmed with an unfiltered search, so
  a hiccup of the site can never produce a false "no longer on the starting list".
- With `use_date_filter = yes` the site is asked only for tournaments ending inside the window. The script verifies
  this while it runs and switches the filter off by itself if the site rejects it. The workflow mode `check-filter`
  compares both ways for three players.

## Storage

The private repo changes only in the small `state/` files. This repo changes almost never (a tiny keep-alive commit
about once a month so that GitHub does not switch off the schedule after 60 days without activity).

## WebSub

The feeds advertise the hubs listed in `config.ini` and the workflow pings them after publishing. A reader that supports
WebSub for third-party feeds can then be notified at once; others simply poll. A dead hub only adds a FAIL line to the
request log.

## Limits

- Only tournaments with a published starting list can be found.
- The script reads an unofficial Excel export; if chess-results changes it, parsing may break.
- RSS items cannot be recalled from a reader once delivered.
- Scheduled GitHub runs can start several minutes late.

## Quick setup

1. Public repo `chess-feed`: upload `chess_watch.py`, `requirements.txt`, `.gitignore`, the READMEs and
   `.github/workflows/watch.yml`. *Settings → Pages → Source: GitHub Actions.*
2. Private repo `chess-watch`: keep only `config.ini`, `players.txt` and `state/`. Delete the old workflow there.
3. Create a fine-grained token with *Contents: Read and write* for the private repo only. In the public repo add the
   Actions secrets `DATA_REPO_TOKEN` (the token) and `DATA_REPO` (`yourname/chess-watch`).
4. *Actions → Update tournaments → Run workflow:* `discover`, optionally `check-filter` and `check-results`, then `update`.
5. Subscribe your reader to a group feed such as `https://YOURNAME.github.io/chess-feed/kikinda.xml`.
