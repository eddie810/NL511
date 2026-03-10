# NL511 Road Conditions — Setup Notes

## GitHub Repo
https://github.com/eddie810/NL511

## KML URL (for Google Earth)
https://raw.githubusercontent.com/eddie810/NL511/main/nl511_roads.kml

## How it works
- GitHub Actions runs nl511_extract.py every hour
- Updates nl511_roads.kml, nl511_data.xlsx, nl511_road_conditions.csv, nl511_events.csv
- Files are committed back to the repo automatically

## To push changes to GitHub
cd "/Users/es/Code/Claude Code/ROADS/NL511"
git add .
git commit -m "your message"
git pull --rebase
git push

## To trigger a manual run
Go to https://github.com/eddie810/NL511/actions
Click NL511 Hourly Update -> Run workflow

## Google Earth Network Link
https://raw.githubusercontent.com/eddie810/NL511/main/nl511_roads.kml
Right-click layer -> Properties -> Refresh to set auto-refresh interval
