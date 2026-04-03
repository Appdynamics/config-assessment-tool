# Troubleshooting

## Docker cannot see input/output/logs

Make sure you mounted the local directories into the container:

```bash
-v <local input dir>:/app/input
-v <local output dir>:/app/output
-v <local logs dir>:/app/logs
```

On Windows, Docker Desktop may also require explicit file sharing for those directories.

## Web UI is not reachable

Check that port `8501` is available and published:

```bash
-p 8501:8501
```

Then open:

```text
http://localhost:8501
```

## macOS blocks the executable bundle

Remove the quarantine attribute once after extracting the bundle:

```bash
sudo xattr -rd com.apple.quarantine .
```

## TLS or certificate verification errors

If the target controller has custom or incomplete certificates, first verify the URL and trust chain. For troubleshooting only, you can temporarily set:

```json
{
  "verifySsl": false
}
```

## Proxy issues

Proxy support has moved here from the main README. Use this section as the canonical reference for proxy setup and troubleshooting.

When `useProxy` is enabled, CAT reads proxy settings from environment variables such as:

```text
HTTP_PROXY
HTTPS_PROXY
WS_PROXY
WSS_PROXY
```

Credentials can also be supplied from `~/.netrc`.

Quick checks:

- Confirm `useProxy` is set to `true` in your job file
- Confirm the proxy environment variable matches your controller protocol (`HTTPS_PROXY` for `https`, `HTTP_PROXY` for `http`)
- Confirm proxy credentials are valid if required by your proxy

## Need more detail in logs

Enable debug logging:

```bash
./config-assessment-tool.sh -d -j DefaultJob
```

or pass `-d` to the backend or Docker command.

## No reports were generated

Check these items:

- the job file exists under `input/jobs/`
- the thresholds file exists under `input/thresholds/`
- output directory is writable
- credentials are valid
- controller host, account, and port are correct

