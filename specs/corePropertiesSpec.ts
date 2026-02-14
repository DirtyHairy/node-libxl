import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import * as xl from '../lib/libxl.js';
import { epsilon } from './testUtils.ts';

describe('The CoreProperties class', () => {
    let book: xl.Book;
    let coreProperties: xl.CoreProperties;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        coreProperties = book.coreProperties();
    });

    it('title and setTitle change the title', () => {
        assert.throws(() => (coreProperties.title as any).call({}));
        assert.throws(() => (coreProperties.setTitle as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setTitle as any).call({}, 'Foo'));

        assert.strictEqual(coreProperties.setTitle('Foo'), coreProperties);
        assert.strictEqual(coreProperties.title(), 'Foo');
    });

    it('subject and setSubject change the title', () => {
        assert.throws(() => (coreProperties.subject as any).call({}));
        assert.throws(() => (coreProperties.setSubject as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setSubject as any).call({}, 'Foo'));

        assert.strictEqual(coreProperties.setSubject('Foo'), coreProperties);
        assert.strictEqual(coreProperties.subject(), 'Foo');
    });

    it('creator and setCreator change the title', () => {
        assert.throws(() => (coreProperties.creator as any).call({}));
        assert.throws(() => (coreProperties.setCreator as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setCreator as any).call({}, 'Foo'));

        assert.strictEqual(coreProperties.setCreator('Foo'), coreProperties);
        assert.strictEqual(coreProperties.creator(), 'Foo');
    });

    it('lastModifiedBy and setLastModifiedBy change the title', () => {
        assert.throws(() => (coreProperties.lastModifiedBy as any).call({}));
        assert.throws(() => (coreProperties.setLastModifiedBy as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setLastModifiedBy as any).call({}, 'Foo'));

        assert.strictEqual(coreProperties.setLastModifiedBy('Foo'), coreProperties);
        assert.strictEqual(coreProperties.lastModifiedBy(), 'Foo');
    });

    it('created and setCreated change the title', () => {
        const now = new Date().toISOString();

        assert.throws(() => (coreProperties.created as any).call({}));
        assert.throws(() => (coreProperties.setCreated as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setCreated as any).call({}, now));

        let date = new Date();
        assert.strictEqual(coreProperties.setCreated(now), coreProperties);
        assert.strictEqual(coreProperties.created(), now);
    });

    it('createdAsDouble and setCreatedAsDouble change the title', () => {
        assert.throws(() => (coreProperties.createdAsDouble as any).call({}));
        assert.throws(() => (coreProperties.setCreatedAsDouble as any).call(coreProperties, 'a'));
        assert.throws(() => (coreProperties.setCreatedAsDouble as any).call({}, 1.0));

        assert.strictEqual(coreProperties.setCreatedAsDouble(1.0), coreProperties);
        assert.ok(coreProperties.createdAsDouble() - 1.0 < epsilon);
    });

    it('modified and setModified change the title', () => {
        const now = new Date().toISOString();

        assert.throws(() => (coreProperties.modified as any).call({}));
        assert.throws(() => (coreProperties.setModified as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setModified as any).call({}, now));

        let date = new Date();
        assert.strictEqual(coreProperties.setModified(now), coreProperties);
        assert.strictEqual(coreProperties.modified(), now);
    });

    it('modifiedAsDouble and setModifiedAsDouble change the title', () => {
        assert.throws(() => (coreProperties.modifiedAsDouble as any).call({}));
        assert.throws(() => (coreProperties.setModifiedAsDouble as any).call(coreProperties, 'a'));
        assert.throws(() => (coreProperties.setModifiedAsDouble as any).call({}, 1.0));

        assert.strictEqual(coreProperties.setModifiedAsDouble(1.0), coreProperties);
        assert.ok(coreProperties.modifiedAsDouble() - 1.0 < epsilon);
    });

    it('tags and setTags change the title', () => {
        assert.throws(() => (coreProperties.tags as any).call({}));
        assert.throws(() => (coreProperties.setTags as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setTags as any).call({}, 'Foo'));

        assert.strictEqual(coreProperties.setTags('Foo'), coreProperties);
        assert.strictEqual(coreProperties.tags(), 'Foo');
    });

    it('categories and setCategories change the title', () => {
        assert.throws(() => (coreProperties.categories as any).call({}));
        assert.throws(() => (coreProperties.setCategories as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setCategories as any).call({}, 'Foo'));

        assert.strictEqual(coreProperties.setCategories('Foo'), coreProperties);
        assert.strictEqual(coreProperties.categories(), 'Foo');
    });

    it('comments and setComments change the title', () => {
        assert.throws(() => (coreProperties.comments as any).call({}));
        assert.throws(() => (coreProperties.setComments as any).call(coreProperties, 1));
        assert.throws(() => (coreProperties.setComments as any).call({}, 'Foo'));

        assert.strictEqual(coreProperties.setComments('Foo'), coreProperties);
        assert.strictEqual(coreProperties.comments(), 'Foo');
    });

    it('removeAll unsets all properties', () => {
        assert.throws(() => (coreProperties.removeAll as any).call({}));

        coreProperties.setTitle('Foo');

        assert.strictEqual(coreProperties.removeAll(), coreProperties);
        assert.throws(() => coreProperties.title());
    });
});
