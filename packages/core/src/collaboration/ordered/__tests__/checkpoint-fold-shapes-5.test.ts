// A fold changes nothing the room uses (design D14): random shape sessions, seeds 89 to 105
// of the fixed list (`sessions/shape-session.ts`).

import { describeShapeSessions } from './sessions/shape-session.ts';

describeShapeSessions([89, 94, 97, 98, 99, 102, 104, 105]);
