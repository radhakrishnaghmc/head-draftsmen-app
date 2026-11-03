import { describe, it, expect } from 'vitest'
import { cleanSignup, validateSignup, type SignupRequest } from '@core/signupRequest'
import { hashPassword, verifyPassword } from '@core/accountPassword'

const ok: SignupRequest = {
  name: 'A B',
  password: 'secret1',
  confirmPassword: 'secret1',
  mobile: '98765 43210',
  email: 'ab@example.com',
  corporation: 'CMC',
  zone: 'Kukatpally',
  circle: ''
}

describe('signup request', () => {
  it('cleans fields', () => {
    expect(cleanSignup({ ...ok, name: '  A   B ', email: ' A.B@Gmail.com ' })).toMatchObject({
      name: 'A B',
      mobile: '9876543210',
      email: 'a.b@gmail.com',
      password: 'secret1'
    })
  })
  it('accepts a complete request (circle optional)', () => {
    expect(validateSignup(ok)).toEqual({})
  })
  it('reports each bad field', () => {
    const e = validateSignup({ ...ok, name: '', password: '123', mobile: '1', email: 'x', corporation: '', zone: '' })
    expect(Object.keys(e).sort()).toEqual(['corporation', 'email', 'mobile', 'name', 'password', 'zone'])
  })
  it('flags mismatched passwords', () => {
    expect(validateSignup({ ...ok, confirmPassword: 'other12' })).toEqual({ confirmPassword: 'The passwords do not match.' })
  })
})

describe('account password', () => {
  it('verifies only the right password and salts each hash', () => {
    const h = hashPassword('secret1')
    expect(verifyPassword('secret1', h)).toBe(true)
    expect(verifyPassword('secret2', h)).toBe(false)
    expect(hashPassword('secret1')).not.toBe(h)
  })
})

describe('account offices', () => {
  const a = { corporation: 'CMC', zone: 'Kukatpally', circle: '' }
  const b = { corporation: 'CMC', zone: 'Quthbullapur', circle: 'Nizampet' }
  const c = { corporation: 'MMC', zone: 'X', circle: '' }

  it('allows at most 2 offices and refuses duplicates', async () => {
    const { addAccountOffice } = await import('@core/signupRequest')
    const two = addAccountOffice([a], b)
    expect(two).toEqual({ offices: [a, b] })
    expect(addAccountOffice([a, b], c)).toMatchObject({ error: expect.stringContaining('Only 2 offices per login') })
    expect(addAccountOffice([a], { ...a })).toMatchObject({ error: expect.stringContaining('already') })
    expect(addAccountOffice([a], { ...b, zone: '' })).toHaveProperty('error')
  })
  it('deleting the sign-up office promotes the second; the last one cannot go', async () => {
    const { removeAccountOffice, addAccountOffice } = await import('@core/signupRequest')
    expect(removeAccountOffice([a, b], a)).toEqual({ offices: [b] })
    expect(removeAccountOffice([b], b)).toHaveProperty('error')
    // freed slot: one more can be added after a deletion
    expect(addAccountOffice([b], c)).toEqual({ offices: [b, c] })
  })
})

describe('google signup details', () => {
  it('needs name, mobile, corporation and zone only', async () => {
    const { validateGoogleSignup } = await import('@core/signupRequest')
    expect(validateGoogleSignup({ name: 'A B', mobile: '9876543210', corporation: 'CMC', zone: 'Kukatpally', circle: '' })).toEqual({})
    expect(Object.keys(validateGoogleSignup({ name: '', mobile: '1', corporation: '', zone: '', circle: '' })).sort()).toEqual(['corporation', 'mobile', 'name', 'zone'])
  })
})
