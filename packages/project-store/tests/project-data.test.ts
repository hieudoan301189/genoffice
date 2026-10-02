import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

// P10: the project-level DVH data lives in the project folder
describe('project data', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'project-store-data-'))
    store = new ProjectStore(tmpDir)
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('round-trips the JSON in <project>/data and goes with the project', () => {
    const project = store.createProject('Cầu Rạch Miễu 2')
    expect(store.readProjectData(project.id)).toBeNull()
    store.writeProjectData(project.id, { format: 'dvh-project-data', objects: [{ id: 'o_1' }] })
    expect(store.readProjectData(project.id)).toEqual({
      format: 'dvh-project-data',
      objects: [{ id: 'o_1' }],
    })
    const file = join(tmpDir, 'projects', project.id, 'data', 'project-data.json')
    expect(existsSync(file)).toBe(true)
    // a damaged file reads as none
    writeFileSync(file, '{ not json')
    expect(store.readProjectData(project.id)).toBeNull()
    store.deleteProject(project.id)
    expect(existsSync(file)).toBe(false)
  })

  it('refuses unknown projects and unsafe ids', () => {
    expect(() => store.writeProjectData('nope', {})).toThrow(/no project/)
    expect(() => store.readProjectData('../escape')).toThrow()
  })
})
