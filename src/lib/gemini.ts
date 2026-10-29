import { ResumeData } from '@/types/resume';

const TECNICS_STANDARD_BENCHMARK = `
TECNICS INTEGRATION TECHNOLOGIES
EXECUTIVE CONSULTANT PROFILE

SECTIONS REQUIRED:
1. Full Name & Target Role
2. Contact Information (Email, Phone, Location, LinkedIn)
3. Professional Summary (Detailed 2-5 sentence executive overview)
4. Technical Skills & Competencies (Grouped by category: Languages, Frameworks, Cloud, Databases, Tools)
5. Professional Experience (Chronological order with Company, Job Title, Dates, Location, and detailed achievement bullet points)
6. Key Projects & Engagements (Project title, technologies used, role, and achievements)
7. Education & Academic Credentials (Degree, Institution, Dates/Year)
8. Professional Certifications
`;

const RESUME_JSON_SCHEMA_EXAMPLE = {
  fullName: "Candidate's Full Name",
  targetRole: "Target Role or Professional Specialization",
  contactInfo: {
    email: "email@example.com",
    phone: "+1 234 567 8900",
    location: "City, State/Country",
    linkedin: "linkedin.com/in/username",
    github: "github.com/username"
  },
  professionalSummary: "A comprehensive executive professional summary detailing years of experience, core domains, and achievements.",
  skills: [
    {
      category: "Operating Systems / Platforms",
      skills: ["Linux", "Windows Server", "Unix"]
    },
    {
      category: "Technologies & Frameworks",
      skills: ["SAP NetWeaver", "S4HANA", "React", "Node.js", "Python"]
    },
    {
      category: "Databases & Cloud",
      skills: ["Oracle", "HANA", "AWS", "Azure"]
    }
  ],
  experience: [
    {
      company: "Enterprise Client or Employer Name",
      role: "Job Title / Consultant Role",
      location: "City, State",
      startDate: "2020",
      endDate: "Present",
      bullets: [
        "Led migration and modernization initiatives with verified metrics.",
        "Architected and deployed enterprise solutions."
      ]
    }
  ],
  projects: [
    {
      name: "Project Name",
      technologies: "Technologies used",
      description: "Project scope and deliverables.",
      bullets: ["Accomplishment 1", "Accomplishment 2"]
    }
  ],
  education: [
    {
      institution: "University / Institute Name",
      degree: "Degree / Qualification Name",
      fieldOfStudy: "Field of Study",
      startDate: "2016",
      endDate: "2020"
    }
  ],
  certifications: [
    "AWS Certified Solutions Architect",
    "SAP Certified Technology Associate"
  ]
};

const FREE_FLASH_MODELS = [
  'gemini-3.5-flash',
  'gemini-flash-lite-latest',
  'gemini-flash-latest'
];

function sanitizeApiKey(key: string): string {
  return key.trim().replace(/^['"]|['"]$/g, '').trim();
}

/**
 * Extracts a clean candidate name from a filename if the resume body lacks one.
 * e.g., "Shekar SAP BASIS Consultant_Tecnics.docx" -> "Shekar"
 * "Vijaya Krishnan G_Standardized.docx" -> "Vijaya Krishnan G"
 */
export function extractNameFromFilename(filename: string): string {
  let clean = filename.replace(/\.[^/.]+$/, ''); // remove extension
  // remove common resume keywords
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

  return { name: 'gemini-3.5-flash', apiVersion: 'v1beta' };
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
    throw new Error(
      'Missing Authorization Key. Please configure your key in Settings or set GEMINI_API_KEY in your environment.'
    );
  }

  const activeKey = sanitizeApiKey(rawKey);

  // If the targetFormatText is a placeholder template (like {{ candidate_name }} {{ content }}),
  // supplement it with the complete standard corporate section guidelines so the AI extracts everything!
  let effectiveBenchmark = targetFormatText;
  if (
    effectiveBenchmark.includes('{{') ||
    effectiveBenchmark.trim().length < 80
  ) {
    effectiveBenchmark = `${TECNICS_STANDARD_BENCHMARK}\n\nTEMPLATE SPECIFICATION:\n${effectiveBenchmark}`;
  }

  const inferredName = filename ? extractNameFromFilename(filename) : '';

  const privacyRule = removePersonalInfo
    ? '6. PRIVACY & ANONYMIZATION: The user explicitly requires all personal contact details (email, phone number, address/location, LinkedIn, GitHub, portfolio) to be REMOVED. Leave contactInfo completely empty: {}. Do not include personal phone numbers, emails, addresses, or profile URLs anywhere in the profile.'
    : '';

  const systemInstruction = `
You are an expert Corporate Resume Specialist at Tecnics Integration Technologies Private Limited.
Your task is to thoroughly analyze the candidate's resume and extract ALL skills, work experiences, projects, education, and achievements into structured JSON.

CRITICAL INSTRUCTIONS:
1. FACTUAL ACCURACY: Extract real facts from the candidate's resume. Do not invent false experience.
2. FULL EXTRACTION: Never omit the candidate's professional summary, skills, work history, or education. If the candidate resume has content, you MUST populate every section.
3. CANDIDATE NAME: Identify the candidate's full name. If the resume does not explicitly state a name, use the inferred name: "${inferredName}".
4. SKILLS GROUPING: Group skills by category (e.g., Operating Systems, Databases, Frameworks, Cloud, Tools).
5. VALID JSON ONLY: Return strictly valid JSON adhering to the schema. Do not wrap in markdown or add conversational text.
${privacyRule}
`;

  const prompt = `
TARGET CORPORATE BENCHMARK:
\"\"\"
${effectiveBenchmark.slice(0, 10000)}
\"\"\"

CANDIDATE SOURCE RESUME:
(File Name: ${filename || 'resume.pdf'})
\"\"\"
${candidateText.slice(0, 25000)}
\"\"\"

Extract and normalize the candidate's complete background into valid JSON matching this schema:
${JSON.stringify(RESUME_JSON_SCHEMA_EXAMPLE, null, 2)}
`;

  let lastError: any = null;

  for (const modelName of FREE_FLASH_MODELS) {
    try {
      console.log(`Executing resume transformation via Free Tier model: ${modelName}`);
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${activeKey}`;

      const requestPayload = {
        contents: [
          {
            role: 'user',
            parts: [{ text: `${systemInstruction}\n\n${prompt}` }],
          },
        ],
        generationConfig: {
          responseMimeType: 'application/json',
          temperature: 0.1,
        },
      };

      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestPayload),
      });

      const resData = await response.json();

      if (!response.ok || resData.error) {
        const errorMsg = resData.error?.message || `HTTP ${response.status}: ${response.statusText}`;
        console.warn(`Model ${modelName} returned error: ${errorMsg}. Trying next model...`);
        lastError = new Error(errorMsg);
        continue;
      }

      const rawCandidateText = resData.candidates?.[0]?.content?.parts?.[0]?.text;

      if (!rawCandidateText) {
        lastError = new Error('Empty response from transformation engine.');
        continue;
      }

      const cleaned = rawCandidateText
        .trim()
        .replace(/^```json\s*/i, '')
        .replace(/^```\s*/i, '')
        .replace(/```\s*$/i, '');

      let parsed: any;
      try {
        parsed = JSON.parse(cleaned);
      } catch (parseErr: any) {
        const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          parsed = JSON.parse(jsonMatch[0]);
        } else {
          throw new Error(`Failed to parse AI output: ${parseErr.message}`);
        }
      }

      // Robust normalization of all fields
      const finalData: ResumeData = {
        fullName:
          (parsed.fullName && parsed.fullName.trim() !== '' && parsed.fullName !== 'Candidate Name')
            ? parsed.fullName.trim()
            : (inferredName || 'Candidate Profile'),
        targetRole: parsed.targetRole || parsed.role || parsed.title || '',
        contactInfo: removePersonalInfo ? {} : (parsed.contactInfo || parsed.contact || {}),
        professionalSummary: removePersonalInfo
          ? scrubPersonalInfo(parsed.professionalSummary || parsed.summary || parsed.profile || '')
          : (parsed.professionalSummary || parsed.summary || parsed.profile || ''),
        skills: normalizeSkills(parsed.skills),
        experience: normalizeExperience(parsed.experience || parsed.workExperience || parsed.work_experience),
        education: normalizeEducation(parsed.education || parsed.academics),
        projects: parsed.projects || parsed.keyProjects || [],
        certifications: normalizeCertifications(parsed.certifications || parsed.licenses),
      };

      return finalData;
    } catch (err: any) {
      lastError = err;
    }
  }

  throw lastError || new Error('Failed to transform resume with available models.');
}

function normalizeSkills(skills: any): any[] {
  if (!skills) return [];
  if (Array.isArray(skills)) {
    // If array of strings ["Java", "Python"]
    if (skills.length > 0 && typeof skills[0] === 'string') {
      return [{ category: 'Core Technical Competencies', skills: skills as string[] }];
    }
    // If array of objects
    return skills.map((item: any) => {
      if (typeof item === 'string') return { category: 'Core', skills: [item] };
      const cat = item.category || item.name || 'Technical Skills';
      const items = Array.isArray(item.skills)
        ? item.skills
        : Array.isArray(item.items)
        ? item.items
        : typeof item.skills === 'string'
        ? item.skills.split(',').map((s: string) => s.trim())
        : [];
      return { category: cat, skills: items };
    });
  }
  return [];
}

function normalizeExperience(exp: any): any[] {
  if (!exp || !Array.isArray(exp)) return [];
  return exp.map((item: any) => ({
    company: item.company || item.companyName || item.employer || item.organization || '',
    role: item.role || item.title || item.jobTitle || item.position || 'Consultant',
    location: item.location || '',
    startDate: item.startDate || item.start || item.duration || '',
    endDate: item.endDate || item.end || '',
    bullets: Array.isArray(item.bullets)
      ? item.bullets
      : Array.isArray(item.responsibilities)
      ? item.responsibilities
      : typeof item.description === 'string'
      ? item.description.split('\n').filter(Boolean)
      : [],
  }));
}

function normalizeEducation(edu: any): any[] {
  if (!edu || !Array.isArray(edu)) return [];
  return edu.map((item: any) => ({
    institution: item.institution || item.school || item.university || item.college || '',
    degree: item.degree || item.qualification || '',
    fieldOfStudy: item.fieldOfStudy || item.major || '',
    startDate: item.startDate || item.start || '',
    endDate: item.endDate || item.end || item.year || item.graduationYear || '',
  }));
}

function normalizeCertifications(certs: any): string[] {
  if (!certs) return [];
  if (Array.isArray(certs)) {
    return certs.map((c: any) => (typeof c === 'string' ? c : c.name || String(c))).filter(Boolean);
  }
  if (typeof certs === 'string') return [certs];
  return [];
}
