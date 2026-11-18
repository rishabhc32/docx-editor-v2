// The item-kind matrix, set `single`, shard 1 of 4 (sharded to keep each process small).

import { matrixShard } from './run.ts';

matrixShard('single', 0, 4);
