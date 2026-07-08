import type { SVGProps } from "react";

function IconBase(props: SVGProps<SVGSVGElement>) {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.8}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      {...props}
    />
  );
}

export function FolderOpenIcon(props: SVGProps<SVGSVGElement>) {
  return (
    <IconBase {...props}>
      <path d="M3 7a2 2 0 0 1 2-2h4l2 2h6a2 2 0 0 1 2 2v1H6.5a2 2 0 0 0-1.94 1.51L3 18V7Z" />
      <path d="M3 18l1.8-6.8A2 2 0 0 1 6.7 9.7H21l-1.9 7.1a2 2 0 0 1-1.93 1.5H5a2 2 0 0 1-2-2Z" />
    </IconBase>
  );
}

export function SaveIcon(props: SVGProps<SVGSVGElement>) {
  return (
    <IconBase {...props}>
      <path d="M5 3h11l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" />
      <path d="M8 3v5h8V3" />
      <path d="M7 21v-7h10v7" />
    </IconBase>
  );
}

export function SaveAsIcon(props: SVGProps<SVGSVGElement>) {
  return (
    <IconBase {...props}>
      <path d="M5 3h9l5 5v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" />
      <path d="M8 3v5h6V3" />
      <path d="M12 13v6M9 16h6" />
    </IconBase>
  );
}

/** 앱 로고 — 스프레드시트 그리드 */
export function AppLogoIcon(props: SVGProps<SVGSVGElement>) {
  return (
    <IconBase {...props}>
      <rect x="3.5" y="3.5" width="17" height="17" rx="2.5" />
      <path d="M3.5 9h17M3.5 14.5h17M9 3.5v17M14.5 3.5v17" />
    </IconBase>
  );
}

/** 업데이트 알림 — 반짝임 */
export function SparkleIcon(props: SVGProps<SVGSVGElement>) {
  return (
    <IconBase {...props}>
      <path d="M12 3.5v3M12 17.5v3M3.5 12h3M17.5 12h3" />
      <path d="M12 8.5 13.4 11.6 16.5 13 13.4 14.4 12 17.5 10.6 14.4 7.5 13 10.6 11.6 12 8.5Z" />
    </IconBase>
  );
}

/** 저장되지 않은 변경 경고 */
export function WarningIcon(props: SVGProps<SVGSVGElement>) {
  return (
    <IconBase {...props}>
      <path d="M12 3.5 21 19.5H3L12 3.5Z" strokeLinejoin="round" />
      <path d="M12 10v4" />
      <path d="M12 16.8v.1" strokeLinecap="round" />
    </IconBase>
  );
}
