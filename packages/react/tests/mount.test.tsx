import { describe } from 'vitest';
import { mountScenarios } from '../../core/tests/unit/mountScenarios';
import { mount } from '../src';

describe('mount() (React)', () => mountScenarios(mount));
