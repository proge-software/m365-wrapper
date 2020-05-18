import { M365Wrapper } from "../src/index";

test("hello", () => {
  const z = new M365Wrapper();
  const q = z.TestStartup();
  expect(q).toEqual(true);
});