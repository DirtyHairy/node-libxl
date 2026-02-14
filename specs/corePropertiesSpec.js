import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import xl from '../lib/libxl.js';
import { epsilon } from './testUtils.js';

describe('The CoreProperties class', () => {
    let book;
    let coreProperties;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        coreProperties = book.coreProperties();
    });

    it('title and setTitle change the title', () => {
        assert.throws(() => coreProperties.title.call({}));
        assert.throws(() => coreProperties.setTitle.call(coreProperties, 1));
        assert.throws(() => coreProperties.setTitle.call({}, 'Foo'));

        assert.strictEqual(coreProperties.setTitle('Foo'), coreProperties);
        assert.strictEqual(coreProperties.title(), 'Foo');
    });

    it('subject and setSubject change the title', () => {
        assert.throws(() => coreProperties.subject.call({}));
        assert.throws(() => coreProperties.setSubject.call(coreProperties, 1));
        assert.throws(() => coreProperties.setSubject.call({}, 'Foo'));

        assert.strictEqual(coreProperties.setSubject('Foo'), coreProperties);
        assert.strictEqual(coreProperties.subject(), 'Foo');
    });

    it('creator and setCreator change the title', () => {
        assert.throws(() => coreProperties.creator.call({}));
        assert.throws(() => coreProperties.setCreator.call(coreProperties, 1));
        assert.throws(() => coreProperties.setCreator.call({}, 'Foo'));

        assert.strictEqual(coreProperties.setCreator('Foo'), coreProperties);
        assert.strictEqual(coreProperties.creator(), 'Foo');
    });

    it('lastModifiedBy and setLastModifiedBy change the title', () => {
        assert.throws(() => coreProperties.lastModifiedBy.call({}));
        assert.throws(() => coreProperties.setLastModifiedBy.call(coreProperties, 1));
        assert.throws(() => coreProperties.setLastModifiedBy.call({}, 'Foo'));

        assert.strictEqual(coreProperties.setLastModifiedBy('Foo'), coreProperties);
        assert.strictEqual(coreProperties.lastModifiedBy(), 'Foo');
    });

    it('created and setCreated change the title', () => {
        const now = new Date().toISOString();

        assert.throws(() => coreProperties.created.call({}));
        assert.throws(() => coreProperties.setCreated.call(coreProperties, 1));
        assert.throws(() => coreProperties.setCreated.call({}, now));

        let date = new Date();
        assert.strictEqual(coreProperties.setCreated(now), coreProperties);
        assert.strictEqual(coreProperties.created(), now);
    });

    it('createdAsDouble and setCreatedAsDouble change the title', () => {
        assert.throws(() => coreProperties.createdAsDouble.call({}));
        assert.throws(() => coreProperties.setCreatedAsDouble.call(coreProperties, 'a'));
        assert.throws(() => coreProperties.setCreatedAsDouble.call({}, 1.0));

        assert.strictEqual(coreProperties.setCreatedAsDouble(1.0), coreProperties);
        assert.ok(coreProperties.createdAsDouble() - 1.0 < epsilon);
    });

    it('modified and setModified change the title', () => {
        const now = new Date().toISOString();

        assert.throws(() => coreProperties.modified.call({}));
        assert.throws(() => coreProperties.setModified.call(coreProperties, 1));
        assert.throws(() => coreProperties.setModified.call({}, now));

        let date = new Date();
        assert.strictEqual(coreProperties.setModified(now), coreProperties);
        assert.strictEqual(coreProperties.modified(), now);
    });

    it('modifiedAsDouble and setModifiedAsDouble change the title', () => {
        assert.throws(() => coreProperties.modifiedAsDouble.call({}));
        assert.throws(() => coreProperties.setModifiedAsDouble.call(coreProperties, 'a'));
        assert.throws(() => coreProperties.setModifiedAsDouble.call({}, 1.0));

        assert.strictEqual(coreProperties.setModifiedAsDouble(1.0), coreProperties);
        assert.ok(coreProperties.modifiedAsDouble() - 1.0 < epsilon);
    });

    it('tags and setTags change the title', () => {
        assert.throws(() => coreProperties.tags.call({}));
        assert.throws(() => coreProperties.setTags.call(coreProperties, 1));
        assert.throws(() => coreProperties.setTags.call({}, 'Foo'));

        assert.strictEqual(coreProperties.setTags('Foo'), coreProperties);
        assert.strictEqual(coreProperties.tags(), 'Foo');
    });

    it('categories and setCategories change the title', () => {
        assert.throws(() => coreProperties.categories.call({}));
        assert.throws(() => coreProperties.setCategories.call(coreProperties, 1));
        assert.throws(() => coreProperties.setCategories.call({}, 'Foo'));

        assert.strictEqual(coreProperties.setCategories('Foo'), coreProperties);
        assert.strictEqual(coreProperties.categories(), 'Foo');
    });

    it('comments and setComments change the title', () => {
        assert.throws(() => coreProperties.comments.call({}));
        assert.throws(() => coreProperties.setComments.call(coreProperties, 1));
        assert.throws(() => coreProperties.setComments.call({}, 'Foo'));

        assert.strictEqual(coreProperties.setComments('Foo'), coreProperties);
        assert.strictEqual(coreProperties.comments(), 'Foo');
    });

    it('removeAll unsets all properties', () => {
        assert.throws(() => coreProperties.removeAll.call({}));

        coreProperties.setTitle('Foo');

        assert.strictEqual(coreProperties.removeAll(), coreProperties);
        assert.throws(() => coreProperties.title());
    });
});
