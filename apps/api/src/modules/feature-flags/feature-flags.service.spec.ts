import { FeatureFlagsService } from './feature-flags.service';

describe('FeatureFlagsService', () => {
  const prismaMock = (flags: Array<{ key: string; enabled: boolean }>) =>
    ({
      featureFlag: {
        findMany: jest.fn().mockResolvedValue(flags),
        upsert: jest.fn().mockResolvedValue({}),
      },
      $transaction: jest.fn().mockImplementation(async (ops: any[]) => {
        // execute sequentially for unit test
        for (const op of ops) await op;
        return [];
      }),
    }) as any;

  const authMock = (isAdmin: boolean) =>
    ({
      isAdminUser: jest.fn().mockResolvedValue(isAdmin),
    }) as any;

  it('admin always sees all modules enabled', async () => {
    const svc = new FeatureFlagsService(
      prismaMock([{ key: 'module_issues', enabled: false }]),
      authMock(true),
    );
    const effective = await svc.getEffectiveForUser(1);
    expect(effective).toEqual({
      module_issues: true,
      module_plans: true,
      module_ai_review: true,
    });
  });

  it('non-admin respects stored flags (missing defaults to true)', async () => {
    const svc = new FeatureFlagsService(
      prismaMock([{ key: 'module_issues', enabled: false }]),
      authMock(false),
    );
    const effective = await svc.getEffectiveForUser(2);
    expect(effective.module_issues).toBe(false);
    expect(effective.module_plans).toBe(true);
    expect(effective.module_ai_review).toBe(true);
  });
});
