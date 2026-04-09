#!/usr/bin/env bash
set -euo pipefail

usage() {
  echo "Usage: $0 <ip_csv> <user> <password> [output_csv]"
  echo
  echo "Example:"
  echo "  $0 devices.csv admin 'secret' uptime-results.csv"
}

trim_spaces() {
  local value="$1"
  # Trim leading whitespace.
  value="${value#"${value%%[![:space:]]*}"}"
  # Trim trailing whitespace.
  value="${value%"${value##*[![:space:]]}"}"
  printf '%s' "$value"
}

csv_escape() {
  local value="$1"
  value="${value//\"/\"\"}"
  printf '"%s"' "$value"
}

if [[ $# -lt 3 || $# -gt 4 ]]; then
  usage
  exit 1
fi

if ! command -v sshpass >/dev/null 2>&1; then
  echo "Error: sshpass is required but not installed."
  echo "Install it and run this script again."
  exit 1
fi

ip_csv="$1"
user="$2"
password="$3"
output_csv="${4:-uptime-$(date +%Y-%m-%d-%H-%M-%S).csv}"

if [[ ! -f "$ip_csv" ]]; then
  echo "Error: IP CSV file not found: $ip_csv"
  exit 1
fi

printf 'ip,uptime,status\n' > "$output_csv"

line_number=0
while IFS=, read -r first_col _rest || [[ -n "${first_col:-}" ]]; do
  line_number=$((line_number + 1))

  ip="${first_col//$'\r'/}"
  ip="${ip#$'\xEF\xBB\xBF'}"   # Strip UTF-8 BOM from first field if present.
  ip="${ip%\"}"
  ip="${ip#\"}"
  ip="$(trim_spaces "$ip")"

  [[ -z "$ip" ]] && continue

  # Allow a simple header row such as "ip".
  if [[ $line_number -eq 1 ]]; then
    header_check="$(printf '%s' "$ip" | tr '[:upper:]' '[:lower:]')"
    if [[ "$header_check" == "ip" || "$header_check" == "host" || "$header_check" == "hostname" ]]; then
      continue
    fi
  fi

  echo "Checking uptime on $ip ..."

  set +e
  uptime_output="$(
    sshpass -p "$password" ssh \
      -o StrictHostKeyChecking=no \
      -o UserKnownHostsFile=/dev/null \
      -o ConnectTimeout=10 \
      "$user@$ip" \
      "uptime -p 2>/dev/null || uptime" 2>&1
  )"
  rc=$?
  set -e

  uptime_output="${uptime_output//$'\n'/ }"

  if [[ $rc -eq 0 ]]; then
    status="ok"
  else
    status="error"
  fi

  printf '%s,%s,%s\n' \
    "$(csv_escape "$ip")" \
    "$(csv_escape "$uptime_output")" \
    "$(csv_escape "$status")" >> "$output_csv"
done < "$ip_csv"

echo "Done. Wrote uptime results to: $output_csv"
