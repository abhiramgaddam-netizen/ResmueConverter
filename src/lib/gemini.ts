import { ResumeData } from '@/types/resume';

export const TECNICS_STANDARD_BENCHMARK = `
TECNICS INTEGRATION TECHNOLOGIES - STANDARD CONSULTANT PROFILE

REQUIRED SECTION ORDER & STRUCTURE:
1. Professional Summary: Bullet list of core experience, years, environments, and OS/cloud familiarity.
2. Specialized Skill Sections (if present in source): Capture domain-specific groupings such as "Migration Projects", "SKILLS@SAP R/3", "Net Weaver Skills", "Hana Skills", or similar technical breakdowns right after Professional Summary.
3. Education: Degree and University (e.g., "MCA from Osmania University.").
4. Technical Skills: Categorized key-value lines (e.g., Operating Systems, SAP Software / Frameworks, Database, Cloud / Tools).
5. Awards and Achievements: Bullet points of client appreciations, awards, and recognitions (if present).
6. Professional Experience: Chronological projects labeled "Project 1", "Project 2", etc., each containing:
   - Client
   - Role
   - Duration
   - Environment (tech stack / versions / tools used in that project)
   - Working Environment (optional bullet points if mentioned in source)
   - Responsibilities (detailed bullet points of tasks and achievements)
`;

export const RESUME_JSON_SCHEMA = {
  type: "OBJECT",
  properties: {
    fullName: { type: "STRING" },
    targetRole: { type: "STRING" },
    professionalSummary: {
      type: "ARRAY",
      items: { type: "STRING" },
      description: "Bullet points summarizing overall experience"
    },
    specializedSkills: {
      type: "ARRAY",
      description: "Domain-specific skill sections or migration project highlights that appear near the top of the resume (e.g., 'Migration Projects', 'SKILLS@SAP R/3', 'Net Weaver Skills', 'Hana Skills')",
      items: {
        type: "OBJECT",
        properties: {
          title: { type: "STRING" },
          bullets: { type: "ARRAY", items: { type: "STRING" } }
        },
        required: ["title", "bullets"]
      }
    },
    education: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          degree: { type: "STRING", description: "Full education line, e.g., 'MCA from Osmania University' or degree name" },
          institution: { type: "STRING" },
          endDate: { type: "STRING" }
        },
        required: ["degree"]
      }
    },
    skills: {
      type: "ARRAY",
      description: "Grouped technical skills (e.g., Operating Systems, SAP Software, Database, Programming Languages)",
      items: {
        type: "OBJECT",
        properties: {
          category: { type: "STRING" },
          skills: { type: "ARRAY", items: { type: "STRING" } }
        },
        required: ["category", "skills"]
      }
    },
    awards: {
      type: "ARRAY",
      items: { type: "STRING" },
      description: "Awards, client appreciations, and achievements"
    },
    experience: {
      type: "ARRAY",
      description: "Professional Experience broken down by Project 1, Project 2, etc.",
      items: {
        type: "OBJECT",
        properties: {
          projectNumber: { type: "STRING", description: "e.g., 'Project 1', 'Project 2'" },
          company: { type: "STRING", description: "Client or Employer Name" },
          role: { type: "STRING", description: "Role or Job Title" },
          startDate: { type: "STRING", description: "Full duration string, e.g., 'September 2015 to till date'" },
          environment: { type: "STRING", description: "Technologies, OS, DBs, and tools used in the project" },
          workingEnvironment: {
            type: "ARRAY",
            items: { type: "STRING" },
            description: "Bullets describing the Working Environment, if explicitly mentioned"
          },
          bullets: {
            type: "ARRAY",
            items: { type: "STRING" },
            description: "All Responsibilities bullet points"
          }
        },
        required: ["projectNumber", "company", "role", "startDate", "bullets"]
      }
    },
    certifications: {
      type: "ARRAY",
      items: { type: "STRING" }
    }
  },
  required: ["fullName", "targetRole", "professionalSummary", "education", "skills", "experience"]
};

export const FREE_FLASH_MODELS = [
  'gemini-3.5-flash',
  'gemini-flash-latest',
  'gemini-3.5-flash-lite',
  'gemini-flash-lite-latest',
  'gemini-3.8-flash'
];

export function sanitizeApiKey(key: string): string {
  return key.trim().replace(/^['"]|['"]$/g, '').trim();
}

export function extractNameFromFilename(filename: string): string {
  let clean = filename.replace(/\.[^/.]+$/, '');
  clean = clean.replace(/(_|\s|-)?(resume|cv|tecnics|standardized|profile|consultant|specialist|developer|engineer|lead|senior|sap|basis).*$/i, '');
  clean = clean.replace(/[_.-]/g, ' ').trim();
  return clean || 'Candidate';
}

export async function resolveAvailableModel(rawKey: string): Promise<{ name: string; apiVersion: string }> {
  const apiKey = sanitizeApiKey(rawKey);

  for (const model of FREE_FLASH_MODELS) {
    try {
      const res = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${model}?key=${apiKey}`
      );
      if (res.ok) {
        return { name: model, apiVersion: 'v1beta' };
      }
    } catch {
      // try next
    }
  }

  return { name: 'gemini-3.8-flash', apiVersion: 'v1beta' };
}

export function scrubPersonalInfo(text: string): string {
  if (!text) return '';
  return text
    // emails
    .replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi, '')
    // phone numbers
    .replace(/(?:\+?\d{1,3}[-.\s]?)?\(?\d{2,4}\)?[-.\s]?\d{3,4}[-.\s]?\d{3,4}/g, '')
    // URLs (LinkedIn, GitHub, websites)
    .replace(/(?:https?:\/\/)?(?:www\.)?(?:linkedin\.com\/[^\s]+|github\.com\/[^\s]+)/gi, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

export async function convertCandidateResume({
  candidateText,
  targetFormatText,
  filename,
  apiKey,
  removePersonalInfo = true,
}: {
  candidateText: string;
  targetFormatText: string;
  filename?: string;
  apiKey?: string;
  removePersonalInfo?: boolean;
}): Promise<ResumeData> {
  const rawKey = apiKey || process.env.GEMINI_API_KEY;
  if (!rawKey) {
    throw new Error('Missing Gemini API Key.');
  }

  const activeKey = sanitizeApiKey(rawKey);
  const inferredName = filename ? extractNameFromFilename(filename) : 'Candidate';

  const systemInstruction = `
You are an expert Corporate Resume Specialist at Tecnics Integration Technologies Private Limited.
Convert the candidate's source resume into the exact Tecnics Standard Consultant format.

CRITICAL RULES:
1. DO NOT DROP SECTIONS: If the candidate has specialized skill sections (like "Migration Projects", "SKILLS@SAP R/3", "Net Weaver Skills", "Hana Skills", etc.), extract ALL of them into 'specializedSkills'.
2. PRESERVE EVERY BULLET: Never summarize or shorten responsibility bullets. Keep all technical details.
3. PROJECT FORMAT: Every work experience or project entry must map to Project 1, Project 2, etc., capturing Client ('company'), Role ('role'), Duration ('startDate'), Environment ('environment'), Working Environment ('workingEnvironment' if present), and Responsibilities ('bullets').
4. CANDIDATE NAME: Extract the candidate's real name, or fallback to "${inferredName}".
5. ZERO FABRICATION: Only extract facts present in the candidate's source text.
${removePersonalInfo ? '6. PRIVACY & ANONYMIZATION: The user explicitly requires personal contact details (phone numbers, emails, addresses, personal URLs) to be omitted.' : ''}
`;

  const userPrompt = `
TARGET BENCHMARK FORMAT:
"""
${TECNICS_STANDARD_BENCHMARK}
"""

CANDIDATE SOURCE RESUME (Filename: ${filename || 'resume.docx'}):
"""
${candidateText.slice(0, 30000)}
"""
`;

  let lastError: any = null;

  for (const modelName of FREE_FLASH_MODELS) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${activeKey}`;

      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: systemInstruction }] },
          contents: [{ role: 'user', parts: [{ text: userPrompt }] }],
          generationConfig: {
            responseMimeType: 'application/json',
            responseSchema: RESUME_JSON_SCHEMA,
            temperature: 0.1,
          },
        }),
        signal: AbortSignal.timeout(25000),
      });

      const resData = await response.json();
      if (!response.ok || resData.error) {
        lastError = new Error(resData.error?.message || `HTTP ${response.status}`);
        continue;
      }

      const rawText = resData.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!rawText) continue;

      const parsed = JSON.parse(rawText);

      // Handle summary if returned as string or array
      let summaryArr: string[] = [];
      if (Array.isArray(parsed.professionalSummary)) {
        summaryArr = parsed.professionalSummary.map((s: any) =>
          removePersonalInfo ? scrubPersonalInfo(String(s)) : String(s)
        );
      } else if (typeof parsed.professionalSummary === 'string' && parsed.professionalSummary.trim()) {
        const text = removePersonalInfo ? scrubPersonalInfo(parsed.professionalSummary) : parsed.professionalSummary;
        summaryArr = text.split('\n').filter(Boolean);
      }

      return {
        fullName:
          parsed.fullName && parsed.fullName.trim() !== '' && parsed.fullName !== 'Candidate'
            ? parsed.fullName.trim()
            : inferredName,
        targetRole: parsed.targetRole || '',
        contactInfo: removePersonalInfo ? {} : (parsed.contactInfo || {}),
        professionalSummary: summaryArr,
        specializedSkills: parsed.specializedSkills || [],
        education: parsed.education || [],
        skills: parsed.skills || [],
        awards: parsed.awards || [],
        experience: (parsed.experience || []).map((exp: any, idx: number) => ({
          projectNumber: exp.projectNumber || `Project ${idx + 1}`,
          company: exp.company || '',
          role: exp.role || '',
          startDate: exp.startDate || '',
          environment: exp.environment || '',
          workingEnvironment: exp.workingEnvironment || [],
          bullets: exp.bullets || [],
        })),
        certifications: parsed.certifications || [],
      };
    } catch (err: any) {
      lastError = err;
    }
  }

  throw lastError || new Error('Failed to transform resume.');
}
