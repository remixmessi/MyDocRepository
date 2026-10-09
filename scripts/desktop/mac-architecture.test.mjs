import {test} from 'node:test';
import assert from 'node:assert/strict';
import {machoArchitectures,assertArchitecture} from './mac-architecture.mjs';
function thin(cpu){const bytes=Buffer.alloc(32);bytes.writeUInt32BE(0xcffaedfe);bytes.writeUInt32LE(cpu,4);return bytes;}
test('thin Mach-O distinguishes Intel and Apple Silicon',()=>{assert.deepEqual(machoArchitectures(thin(0x01000007)),['x64']);assert.deepEqual(machoArchitectures(thin(0x0100000c)),['arm64']);assert.deepEqual(machoArchitectures(Buffer.from('ELF')),[]);});
test('universal Mach-O and malformed architecture tables',()=>{const bytes=Buffer.alloc(48);bytes.writeUInt32BE(0xcafebabe);bytes.writeUInt32BE(2,4);bytes.writeUInt32BE(0x01000007,8);bytes.writeUInt32BE(0x0100000c,28);assert.deepEqual(machoArchitectures(bytes),['x64','arm64']);assert.throws(()=>machoArchitectures(bytes.subarray(0,28)),/无效/);});
test('reject arm64 runtime mislabeled x64 before producing installers',()=>{assert.throws(()=>assertArchitecture(['arm64'],['x64'],'MyDoc.app'),/目标 x64，实际 arm64/);assert.doesNotThrow(()=>assertArchitecture(['x64'],['x64'],'MyDoc.app'));assert.throws(()=>assertArchitecture(['x64'],['x64','arm64'],'MyDoc.app'),/架构不匹配/);});
