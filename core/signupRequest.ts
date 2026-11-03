/** A new account, filled in on the login screen's Sign up page. The user picks their own Login ID and password. */
export interface SignupRequest {
  name: string
  password: string
  confirmPassword: string
  mobile: string
  /** Doubles as the Login ID. Must be confirmed by clicking the link emailed to it before the account is created. */
  email: string
  corporation: string
  zone: string
  /** Empty for a zone-level office. */
  circle: string
}

export type SignupErrors = Partial<Record<keyof SignupRequest, string>>

const EMAIL = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/

export const MIN_PASSWORD_LENGTH = 6

/** Trims every field and lower-cases the Login ID and email, so what is saved is what was meant. The password is left exactly as typed. */
export function cleanSignup(r: SignupRequest): SignupRequest {
  return {
    name: r.name.trim().replace(/\s+/g, ' '),
    password: r.password,
    confirmPassword: r.confirmPassword,
    mobile: r.mobile.replace(/[^\d+]/g, ''),
    email: r.email.trim().toLowerCase(),
    corporation: r.corporation.trim(),
    zone: r.zone.trim(),
    circle: r.circle.trim()
  }
}

/** Returns the problems with a request (empty object when it is fine). Works on the cleaned values. */
export function validateSignup(raw: SignupRequest): SignupErrors {
  const r = cleanSignup(raw)
  const errors: SignupErrors = {}
  if (r.name.length < 2) errors.name = 'Enter your name.'
  if (r.password.length < MIN_PASSWORD_LENGTH) errors.password = `Use at least ${MIN_PASSWORD_LENGTH} characters.`
  else if (r.password !== r.confirmPassword) errors.confirmPassword = 'The passwords do not match.'
  if (r.mobile.replace(/\D/g, '').length < 10) errors.mobile = 'Enter a 10-digit mobile number.'
  if (!EMAIL.test(r.email)) errors.email = 'Enter a valid email address.'
  if (!r.corporation) errors.corporation = 'Choose your corporation.'
  if (!r.zone) errors.zone = 'Choose your zone.'
  return errors
}

/** An office on a self-registered account: Corporation → Zone → Circle ("" circle = zonal office). */
export interface AccountOffice {
  corporation: string
  zone: string
  circle: string
}

/** A self-registered Login ID may hold at most this many offices. */
export const MAX_ACCOUNT_OFFICES = 2

export const sameAccountOffice = (a: AccountOffice, b: AccountOffice): boolean =>
  a.corporation.trim().toLowerCase() === b.corporation.trim().toLowerCase() &&
  a.zone.trim().toLowerCase() === b.zone.trim().toLowerCase() &&
  a.circle.trim().toLowerCase() === b.circle.trim().toLowerCase()

/** Adds an office to an account's list; the reason it cannot be added otherwise. */
export function addAccountOffice(
  offices: AccountOffice[],
  office: AccountOffice
): { offices: AccountOffice[] } | { error: string } {
  const next: AccountOffice = { corporation: office.corporation.trim(), zone: office.zone.trim(), circle: office.circle.trim() }
  if (!next.corporation || !next.zone) return { error: 'Choose the corporation and zone of the office.' }
  if (offices.length >= MAX_ACCOUNT_OFFICES) {
    return { error: `Only ${MAX_ACCOUNT_OFFICES} offices per login are allowed. To add another, delete one of the two.` }
  }
  if (offices.some((o) => sameAccountOffice(o, next))) return { error: 'That office is already on your login.' }
  return { offices: [...offices, next] }
}

/** Removes an office; the first remaining one becomes the account's primary (sign-up) office. An account always keeps one. */
export function removeAccountOffice(
  offices: AccountOffice[],
  office: AccountOffice
): { offices: AccountOffice[] } | { error: string } {
  if (!offices.some((o) => sameAccountOffice(o, office))) return { error: 'That office is not on your login.' }
  if (offices.length <= 1) return { error: 'A login needs at least one office. Add another before deleting this one.' }
  return { offices: offices.filter((o) => !sameAccountOffice(o, office)) }
}

/** The details a first-time Google user adds before their account is created (email and password come from Google). */
export type GoogleSignupDetails = Pick<SignupRequest, 'name' | 'mobile' | 'corporation' | 'zone' | 'circle'>

export function validateGoogleSignup(raw: GoogleSignupDetails): SignupErrors {
  const r = cleanSignup({ ...raw, password: '', confirmPassword: '', email: '' })
  const errors: SignupErrors = {}
  if (r.name.length < 2) errors.name = 'Enter your name.'
  if (r.mobile.replace(/\D/g, '').length < 10) errors.mobile = 'Enter a 10-digit mobile number.'
  if (!r.corporation) errors.corporation = 'Choose your corporation.'
  if (!r.zone) errors.zone = 'Choose your zone.'
  return errors
}
