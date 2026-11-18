// People for tests that name a person by a short name: the ID is the name in lower case, and
// there is no company, so the Word name (`w:author`) is the name itself.

import { personAttribution, type OrderedPerson } from '../model/person.ts';
import type { OrderedAttribution } from '../model/types.ts';

/** The person a test names `name`. */
export function testPerson(name: string): OrderedPerson {
  return { id: name.toLowerCase(), name, org: '' };
}

/** What `name` makes, as the room stamps it, with a date or none. */
export function testAttribution(name: string, date?: string): OrderedAttribution {
  if (date === undefined) {
    return { author: name, person: testPerson(name) };
  }

  return personAttribution(testPerson(name), date);
}
