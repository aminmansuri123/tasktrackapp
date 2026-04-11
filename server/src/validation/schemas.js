const { z } = require('zod');

/** Workspace JSON from client: known keys only; unknown keys stripped. Array caps limit abuse while allowing large real workspaces. */
const workspacePutSchema = z
  .object({
    users: z.array(z.unknown()).max(5000).optional(),
    tasks: z.array(z.unknown()).max(200000).optional(),
    locations: z.array(z.unknown()).max(5000).optional(),
    segregationTypes: z.array(z.unknown()).max(1000).optional(),
    holidays: z.array(z.unknown()).max(10000).optional(),
    notes: z.array(z.unknown()).max(50000).optional(),
    learningNotes: z.array(z.unknown()).max(50000).optional(),
    milestones: z.array(z.unknown()).max(20000).optional(),
    dailyPlanner: z.array(z.unknown()).max(200000).optional(),
    locationItems: z.array(z.unknown()).max(300000).optional(),
    codeSnippets: z.array(z.unknown()).max(20000).optional(),
    journal: z.record(z.string(), z.unknown()).optional(),
    reportToOptions: z.array(z.unknown()).max(200).optional(),
    templateBlocks: z.array(z.unknown()).max(5000).optional(),
    notificationEmailCc: z.array(z.string().max(320)).max(50).optional(),
  });

const loginBodySchema = z.object({
  email: z.string().trim().min(1).max(320),
  password: z.string().min(1).max(500),
});

const registerBodySchema = z.object({
  name: z.string().trim().min(1).max(200),
  email: z.string().trim().email().max(320),
  password: z.string().min(6).max(500),
  accountType: z.enum(['team_user', 'org_admin']).optional(),
  orgAdminUserId: z.union([z.number().int().positive(), z.string()]).optional(),
});

const forgotPasswordRequestSchema = z.object({
  email: z.string().trim().email().max(320),
});

const forgotPasswordResetSchema = z.object({
  email: z.string().trim().email().max(320),
  code: z.union([z.string(), z.number()]),
  newPassword: z.string().min(6).max(500),
});

const changePasswordSchema = z.object({
  currentPassword: z.string().min(1).max(500),
  newPassword: z.string().min(6).max(500),
});

const requestApprovalSchema = z.object({
  name: z.string().trim().min(1).max(200),
  email: z.string().trim().email().max(320),
});

const emailTaskViewSummarySchema = z.object({
  context: z
    .object({
      tileKey: z.string().trim().max(80).optional(),
      tileLabel: z.string().trim().max(120).optional(),
    })
    .optional(),
  recipients: z
    .array(
      z.object({
        userId: z.preprocess((v) => {
          if (typeof v === 'string') return parseInt(v, 10);
          return v;
        }, z.number().int().positive()),
        tasks: z
          .array(
            z.object({
              title: z.string().max(2000),
              due: z.string().max(120).optional(),
              overdue: z.boolean().optional(),
              status: z.string().max(80).optional(),
            })
          )
          .max(5000),
      })
    )
    .max(200),
});

/**
 * Validate backup restore body: either a workspace object or { data: workspace }.
 * @returns {{ ok: true, data: object } | { ok: false, details: object }}
 */
function parseWorkspaceRestoreBody(body) {
  const raw = body && typeof body === 'object' ? body : {};
  const candidate = raw.data !== undefined && raw.data !== null ? raw.data : raw;
  const parsed = workspacePutSchema.safeParse(candidate);
  if (!parsed.success) {
    return { ok: false, details: parsed.error.flatten() };
  }
  return { ok: true, data: parsed.data };
}

module.exports = {
  workspacePutSchema,
  loginBodySchema,
  registerBodySchema,
  forgotPasswordRequestSchema,
  forgotPasswordResetSchema,
  changePasswordSchema,
  requestApprovalSchema,
  emailTaskViewSummarySchema,
  parseWorkspaceRestoreBody,
};
