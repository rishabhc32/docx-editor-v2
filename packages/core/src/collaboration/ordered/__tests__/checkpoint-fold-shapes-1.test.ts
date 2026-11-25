// A fold changes nothing the room uses (design D14): random shape sessions, seeds 5 to 22
// of the fixed list (`sessions/shape-session.ts`).

import { describeShapeSessions } from './sessions/shape-session.ts';

describeShapeSessions([5, 6, 12, 14, 15, 16, 20, 22]);
