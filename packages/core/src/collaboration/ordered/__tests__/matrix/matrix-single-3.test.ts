// The item-kind matrix, set `single`, shard 3 of 4 (sharded to keep each process small).

import { matrixShard } from './run.ts';

matrixShard('single', 2, 4);
