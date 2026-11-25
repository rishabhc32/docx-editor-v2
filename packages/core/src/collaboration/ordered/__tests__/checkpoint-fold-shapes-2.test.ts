// A fold changes nothing the room uses (design D14): random shape sessions, seeds 23 to 33
// of the fixed list (`sessions/shape-session.ts`).

import { describeShapeSessions } from './sessions/shape-session.ts';

describeShapeSessions([23, 25, 26, 27, 28, 31, 32, 33]);
