const GITHUB_OWNER = "baggychani";
const GITHUB_REPO = "GichanExcel";
const GITHUB_API_LATEST_RELEASE = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/releases/latest`;
export const GITHUB_RELEASE_URL = `https://github.com/${GITHUB_OWNER}/${GITHUB_REPO}/releases`;

const LAST_CHECK_KEY = "gichan-excel:last-update-check";
const CHECK_INTERVAL_MS = 12 * 60 * 60 * 1000;

export interface UpdateInfo {
  version: string;
  url: string;
  notes: string;
}

interface GithubReleaseResponse {
  tag_name?: string;
  html_url?: string;
  body?: string;
  draft?: boolean;
  prerelease?: boolean;
}

function shouldCheckNow(): boolean {
  const lastCheck = localStorage.getItem(LAST_CHECK_KEY);
  if (!lastCheck) {
    return true;
  }
  const lastCheckTime = Date.parse(lastCheck);
  if (Number.isNaN(lastCheckTime)) {
    return true;
  }
  return Date.now() - lastCheckTime >= CHECK_INTERVAL_MS;
}

function markCheckedNow(): void {
  try {
    localStorage.setItem(LAST_CHECK_KEY, new Date().toISOString());
  } catch {
    // localStorage를 못 쓰는 환경이면 매번 새로 확인하게 되는 정도라 무시해도 안전합니다.
  }
}

/** "1.2.3" 형태의 버전 문자열을 비교합니다. latest가 current보다 높으면 true. */
function isNewerVersion(current: string, latest: string): boolean {
  const parse = (value: string) =>
    value
      .split(".")
      .map((part) => Number.parseInt(part, 10))
      .map((part) => (Number.isFinite(part) ? part : 0));

  const currentParts = parse(current);
  const latestParts = parse(latest);
  const length = Math.max(currentParts.length, latestParts.length);

  for (let i = 0; i < length; i += 1) {
    const c = currentParts[i] ?? 0;
    const l = latestParts[i] ?? 0;
    if (l > c) return true;
    if (l < c) return false;
  }
  return false;
}

/**
 * GitHub의 최신 릴리스를 확인해 현재 버전보다 높은 버전이 있으면 정보를 반환합니다.
 * 마지막 확인 후 12시간이 지나지 않았다면 네트워크 요청 없이 즉시 null을 반환합니다.
 * (요청이 실패하더라도 "확인 시도" 자체는 기록해 반복 재시도를 방지합니다.)
 */
export async function checkForUpdate(
  currentVersion: string,
): Promise<UpdateInfo | null> {
  if (!shouldCheckNow()) {
    return null;
  }

  try {
    const response = await fetch(GITHUB_API_LATEST_RELEASE, {
      headers: { Accept: "application/vnd.github+json" },
    });

    if (!response.ok) {
      return null;
    }

    const data = (await response.json()) as GithubReleaseResponse;
    if (data.draft || data.prerelease) {
      return null;
    }

    const latestTag = (data.tag_name ?? "").trim();
    const latestVersion = latestTag.replace(/^v/i, "");

    if (!latestVersion || !isNewerVersion(currentVersion, latestVersion)) {
      return null;
    }

    return {
      version: latestVersion,
      url: data.html_url ?? GITHUB_RELEASE_URL,
      notes: (data.body ?? "").trim(),
    };
  } catch {
    return null;
  } finally {
    markCheckedNow();
  }
}
