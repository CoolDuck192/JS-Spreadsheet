import { useLayoutEffect, useRef, useState } from "react";

const MENU_VIEWPORT_MARGIN = 8;

export function useClampedMenuPosition(x: number, y: number) {
  const menuRef = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: x, top: y });

  useLayoutEffect(() => {
    const menu = menuRef.current;
    if (!menu) {
      setPosition({ left: x, top: y });
      return;
    }

    function updatePosition() {
      const rect = menu.getBoundingClientRect();
      const maxLeft = Math.max(
        MENU_VIEWPORT_MARGIN,
        window.innerWidth - rect.width - MENU_VIEWPORT_MARGIN
      );
      const maxTop = Math.max(
        MENU_VIEWPORT_MARGIN,
        window.innerHeight - rect.height - MENU_VIEWPORT_MARGIN
      );
      const nextPosition = {
        left: Math.min(Math.max(MENU_VIEWPORT_MARGIN, x), maxLeft),
        top: Math.min(Math.max(MENU_VIEWPORT_MARGIN, y), maxTop)
      };
      setPosition((currentPosition) =>
        currentPosition.left === nextPosition.left && currentPosition.top === nextPosition.top
          ? currentPosition
          : nextPosition
      );
    }

    updatePosition();
    window.addEventListener("resize", updatePosition);
    return () => window.removeEventListener("resize", updatePosition);
  }, [x, y]);

  return { menuRef, position };
}
