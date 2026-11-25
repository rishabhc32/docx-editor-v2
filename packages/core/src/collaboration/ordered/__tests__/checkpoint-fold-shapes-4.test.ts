// A fold changes nothing the room uses (design D14): random shape sessions, seeds 65 to 86
// of the fixed list (`sessions/shape-session.ts`).

import { describeShapeSessions } from './sessions/shape-session.ts';

describeShapeSessions([65, 66, 69, 76, 79, 80, 82, 86]);
