// A fold changes nothing the room uses (design D14): random shape sessions, seeds 107 to 118
// of the fixed list (`sessions/shape-session.ts`).

import { describeShapeSessions } from './sessions/shape-session.ts';

describeShapeSessions([107, 109, 110, 112, 113, 117, 118]);
