// A fold changes nothing the room uses (design D14): random shape sessions, seeds 35 to 61
// of the fixed list (`sessions/shape-session.ts`).

import { describeShapeSessions } from './sessions/shape-session.ts';

describeShapeSessions([35, 38, 41, 45, 46, 48, 51, 61]);
